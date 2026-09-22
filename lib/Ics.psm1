<#
    Ics - just enough iCalendar to answer "what is on my calendar today?".

    Published Outlook feeds are mostly recurring meetings, so a parser that
    ignores RRULE shows an almost empty day. This one expands recurrences for a
    single date: FREQ DAILY/WEEKLY/MONTHLY/YEARLY with INTERVAL, BYDAY,
    BYMONTHDAY, BYMONTH, BYSETPOS, WKST, COUNT and UNTIL, plus RDATE, EXDATE and
    the RECURRENCE-ID overrides Outlook writes when a single occurrence moves.

    Get-IcsEvents -Text <ics> -Date <day>  ->  Start End Subject Location AllDay Busy Oof
#>

$ErrorActionPreference = 'Stop'

$script:DOW = @{ SU = 0; MO = 1; TU = 2; WE = 3; TH = 4; FR = 5; SA = 6 }

# Outlook writes Windows time-zone ids, which .NET knows. Feeds from elsewhere
# use IANA names, which Windows PowerShell does not - translate the common ones
# and fall back to local time for anything else.
$script:TZ_ALIAS = @{
    'UTC'                  = 'UTC'
    'Etc/UTC'              = 'UTC'
    'Etc/GMT'              = 'UTC'
    'Europe/Berlin'        = 'W. Europe Standard Time'
    'Europe/Vienna'        = 'W. Europe Standard Time'
    'Europe/Zurich'        = 'W. Europe Standard Time'
    'Europe/Rome'          = 'W. Europe Standard Time'
    'Europe/Amsterdam'     = 'W. Europe Standard Time'
    'Europe/Stockholm'     = 'W. Europe Standard Time'
    'Europe/Oslo'          = 'W. Europe Standard Time'
    'Europe/Copenhagen'    = 'Romance Standard Time'
    'Europe/Paris'         = 'Romance Standard Time'
    'Europe/Madrid'        = 'Romance Standard Time'
    'Europe/Brussels'      = 'Romance Standard Time'
    'Europe/Warsaw'        = 'Central European Standard Time'
    'Europe/Prague'        = 'Central Europe Standard Time'
    'Europe/Budapest'      = 'Central Europe Standard Time'
    'Europe/London'        = 'GMT Standard Time'
    'Europe/Dublin'        = 'GMT Standard Time'
    'Europe/Lisbon'        = 'GMT Standard Time'
    'Europe/Helsinki'      = 'FLE Standard Time'
    'Europe/Kiev'          = 'FLE Standard Time'
    'Europe/Athens'        = 'GTB Standard Time'
    'Europe/Bucharest'     = 'GTB Standard Time'
    'Europe/Istanbul'      = 'Turkey Standard Time'
    'Europe/Moscow'        = 'Russian Standard Time'
    'America/New_York'     = 'Eastern Standard Time'
    'America/Toronto'      = 'Eastern Standard Time'
    'America/Chicago'      = 'Central Standard Time'
    'America/Denver'       = 'Mountain Standard Time'
    'America/Los_Angeles'  = 'Pacific Standard Time'
    'America/Sao_Paulo'    = 'E. South America Standard Time'
    'Asia/Jerusalem'       = 'Israel Standard Time'
    'Asia/Dubai'           = 'Arabian Standard Time'
    'Asia/Kolkata'         = 'India Standard Time'
    'Asia/Calcutta'        = 'India Standard Time'
    'Asia/Shanghai'        = 'China Standard Time'
    'Asia/Singapore'       = 'Singapore Standard Time'
    'Asia/Tokyo'           = 'Tokyo Standard Time'
    'Asia/Seoul'           = 'Korea Standard Time'
    'Australia/Sydney'     = 'AUS Eastern Standard Time'
    'Australia/Melbourne'  = 'AUS Eastern Standard Time'
    'Pacific/Auckland'     = 'New Zealand Standard Time'
}
$script:tzCache = @{}

function Get-IcsTimeZone([string]$Id) {
    if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
    $key = $Id.Trim('"')
    if ($script:tzCache.ContainsKey($key)) { return $script:tzCache[$key] }
    $tz = $null
    foreach ($cand in @($key, $script:TZ_ALIAS[$key])) {
        if (-not $cand) { continue }
        try { $tz = [TimeZoneInfo]::FindSystemTimeZoneById($cand); break } catch { }
    }
    $script:tzCache[$key] = $tz
    $tz
}

# --------------------------------------------------------------------------
# lexing
# --------------------------------------------------------------------------
function Expand-IcsFolding([string]$Text) {
    # RFC 5545 folds long lines: a CRLF followed by one space or tab continues
    # the line and the whitespace is not part of the value.
    ($Text -replace "\r\n[ \t]", '') -replace "[\r\n]+[ \t]", ''
}

function ConvertFrom-IcsValue([string]$Value) {
    # \n \N -> newline, \, \; \\ -> the bare character
    $v = $Value -replace '\\[nN]', "`n"
    $v -replace '\\([,;\\])', '$1'
}

function Split-IcsLine([string]$Line) {
    # NAME;PARAM=a;PARAM2="b:c":VALUE   - the colon inside quotes is not the separator
    $inQuotes = $false
    $colon    = -1
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $c = $Line[$i]
        if     ($c -eq '"') { $inQuotes = -not $inQuotes }
        elseif ($c -eq ':' -and -not $inQuotes) { $colon = $i; break }
    }
    if ($colon -lt 0) { return $null }

    $head  = $Line.Substring(0, $colon)
    $value = $Line.Substring($colon + 1)

    $parts  = @()
    $buf    = New-Object Text.StringBuilder
    $inQ    = $false
    foreach ($c in $head.ToCharArray()) {
        if     ($c -eq '"') { $inQ = -not $inQ; [void]$buf.Append($c) }
        elseif ($c -eq ';' -and -not $inQ) { $parts += $buf.ToString(); [void]$buf.Clear() }
        else   { [void]$buf.Append($c) }
    }
    $parts += $buf.ToString()

    $params = @{}
    foreach ($p in $parts[1..($parts.Count - 1)]) {
        $eq = $p.IndexOf('=')
        if ($eq -gt 0) { $params[$p.Substring(0, $eq).ToUpperInvariant()] = $p.Substring($eq + 1).Trim('"') }
    }

    @{ Name = $parts[0].ToUpperInvariant(); Params = $params; Value = $value }
}

function ConvertFrom-IcsText([string]$Text) {
<#
    .SYNOPSIS Every VEVENT as a hashtable of property name -> @{Value;Params}.
              Properties that may repeat (EXDATE, RDATE) collect into a list.
#>
    $events = @()
    $cur    = $null
    $depth  = 0     # so a VALARM inside a VEVENT does not leak its properties

    foreach ($line in (Expand-IcsFolding $Text) -split "`r?`n") {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $p = Split-IcsLine $line.TrimEnd()
        if (-not $p) { continue }

        if ($p.Name -eq 'BEGIN') {
            if ($p.Value.ToUpperInvariant() -eq 'VEVENT') { $cur = @{}; $depth = 0 }
            elseif ($null -ne $cur) { $depth++ }
            continue
        }
        if ($p.Name -eq 'END') {
            if ($p.Value.ToUpperInvariant() -eq 'VEVENT') {
                if ($null -ne $cur) { $events += ,$cur }
                $cur = $null
            } elseif ($null -ne $cur -and $depth -gt 0) { $depth-- }
            continue
        }
        if ($null -eq $cur -or $depth -gt 0) { continue }

        $entry = @{ Value = $p.Value; Params = $p.Params }
        if ($p.Name -in @('EXDATE', 'RDATE')) {
            if (-not $cur.ContainsKey($p.Name)) { $cur[$p.Name] = @() }
            $cur[$p.Name] += ,$entry
        } elseif (-not $cur.ContainsKey($p.Name)) {
            $cur[$p.Name] = $entry
        }
    }
    # plain output on purpose: ",$events" would survive @() at the call site as a
    # single element holding the whole array
    $events
}

# --------------------------------------------------------------------------
# dates
# --------------------------------------------------------------------------
function ConvertFrom-IcsDate([string]$Value, [hashtable]$Params) {
<#
    .SYNOPSIS One DATE or DATE-TIME as local time. Returns @{At; AllDay}.
#>
    $v = $Value.Trim()
    if (-not $v) { return $null }

    $allDay = ($Params -and $Params['VALUE'] -eq 'DATE') -or ($v.Length -eq 8)
    if ($allDay) {
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($v.Substring(0, 8), 'yyyyMMdd', [cultureinfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::None, [ref]$d)) {
            return @{ At = $d; AllDay = $true }
        }
        return $null
    }

    $utc = $v.EndsWith('Z')
    $raw = if ($utc) { $v.Substring(0, $v.Length - 1) } else { $v }
    $dt  = [datetime]::MinValue
    $ok  = $false
    # one format at a time: PowerShell folds a format array into a single string
    # and then silently matches nothing
    foreach ($fmt in @('yyyyMMdd\THHmmss', 'yyyyMMdd\THHmm')) {
        if ([datetime]::TryParseExact($raw, $fmt, [cultureinfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::None, [ref]$dt)) { $ok = $true; break }
    }
    if (-not $ok) { return $null }

    if ($utc) {
        $local = [datetime]::SpecifyKind($dt, [DateTimeKind]::Utc).ToLocalTime()
        return @{ At = $local; AllDay = $false }
    }

    $tz = if ($Params) { Get-IcsTimeZone $Params['TZID'] } else { $null }
    if ($tz) {
        $u = [datetime]::SpecifyKind($dt, [DateTimeKind]::Unspecified)
        return @{ At = [TimeZoneInfo]::ConvertTime($u, $tz, [TimeZoneInfo]::Local); AllDay = $false }
    }
    # floating time, or a zone this machine does not know: read it as local
    @{ At = $dt; AllDay = $false }
}

function Get-IcsProperty($Event, [string]$Name) {
    if ($Event.ContainsKey($Name)) { ConvertFrom-IcsValue $Event[$Name].Value } else { $null }
}

function Get-IcsDateList($Entries) {
    # EXDATE / RDATE - each line may carry a comma separated list
    $out = @()
    foreach ($e in @($Entries)) {
        if (-not $e) { continue }
        foreach ($piece in ($e.Value -split ',')) {
            $d = ConvertFrom-IcsDate $piece $e.Params
            if ($d) { $out += $d.At }
        }
    }
    $out
}

function ConvertFrom-IcsDuration([string]$Text) {
    # P1DT2H30M -> TimeSpan
    if ($Text -notmatch '^(?<s>[+-])?P(?:(?<w>\d+)W)?(?:(?<d>\d+)D)?(?:T(?:(?<h>\d+)H)?(?:(?<m>\d+)M)?(?:(?<sec>\d+)S)?)?$') { return $null }
    $ts = [timespan]::FromDays(([int]$matches['w']) * 7 + [int]$matches['d']) +
          [timespan]::FromHours([int]$matches['h']) +
          [timespan]::FromMinutes([int]$matches['m']) +
          [timespan]::FromSeconds([int]$matches['sec'])
    if ($matches['s'] -eq '-') { return $ts.Negate() }
    $ts
}

# --------------------------------------------------------------------------
# recurrence
# --------------------------------------------------------------------------
function ConvertFrom-IcsRule([string]$Text) {
    $r = @{}
    foreach ($part in ($Text -split ';')) {
        $eq = $part.IndexOf('=')
        if ($eq -gt 0) { $r[$part.Substring(0, $eq).ToUpperInvariant()] = $part.Substring($eq + 1) }
    }
    $r
}

function Get-WeekStart([datetime]$Date, [int]$WkSt) {
    $Date.Date.AddDays(-((([int]$Date.DayOfWeek) - $WkSt + 7) % 7))
}

function Get-MonthDays([datetime]$Month, $Rule, [datetime]$DtStart) {
<#
    .SYNOPSIS Which days of $Month the rule selects (day-of-month numbers).
#>
    $first = Get-Date -Year $Month.Year -Month $Month.Month -Day 1
    $len   = [datetime]::DaysInMonth($Month.Year, $Month.Month)

    if ($Rule['BYDAY']) {
        $wanted = @{}
        $ordinals = @{}
        foreach ($tok in ($Rule['BYDAY'] -split ',')) {
            if ($tok -match '^([+-]?\d+)?(SU|MO|TU|WE|TH|FR|SA)$') {
                $d = $script:DOW[$matches[2]]
                $wanted[$d] = $true
                if ($matches[1]) {
                    if (-not $ordinals.ContainsKey($d)) { $ordinals[$d] = @() }
                    $ordinals[$d] += [int]$matches[1]
                }
            }
        }
        # every matching day in the month, in order
        $hits = @()
        for ($i = 1; $i -le $len; $i++) {
            $d = $first.AddDays($i - 1)
            if ($wanted[[int]$d.DayOfWeek]) { $hits += $d }
        }
        if ($ordinals.Count -gt 0) {
            $keep = @()
            foreach ($dow in $ordinals.Keys) {
                $ofDow = @($hits | Where-Object { [int]$_.DayOfWeek -eq $dow })
                foreach ($n in $ordinals[$dow]) {
                    $idx = if ($n -gt 0) { $n - 1 } else { $ofDow.Count + $n }
                    if ($idx -ge 0 -and $idx -lt $ofDow.Count) { $keep += $ofDow[$idx] }
                }
            }
            $hits = @($keep | Sort-Object)
        }
        if ($Rule['BYSETPOS']) {
            $keep = @()
            foreach ($n in ($Rule['BYSETPOS'] -split ',')) {
                $i = [int]$n
                $idx = if ($i -gt 0) { $i - 1 } else { $hits.Count + $i }
                if ($idx -ge 0 -and $idx -lt $hits.Count) { $keep += $hits[$idx] }
            }
            $hits = @($keep | Sort-Object)
        }
        return @($hits | ForEach-Object { $_.Day })
    }

    if ($Rule['BYMONTHDAY']) {
        $days = @()
        foreach ($n in ($Rule['BYMONTHDAY'] -split ',')) {
            $i = [int]$n
            $d = if ($i -gt 0) { $i } else { $len + $i + 1 }
            if ($d -ge 1 -and $d -le $len) { $days += $d }
        }
        return $days
    }

    # no BY* part: the rule repeats the start's day of month
    if ($DtStart.Day -le $len) { return @($DtStart.Day) } else { return @() }
}

function Test-IcsRuleDay($Rule, [datetime]$DtStart, [datetime]$Day) {
<#
    .SYNOPSIS Does an occurrence of this rule start on $Day? (UNTIL/COUNT aside)
#>
    $day = $Day.Date
    $s   = $DtStart.Date
    if ($day -lt $s) { return $false }

    $interval = if ($Rule['INTERVAL']) { [int]$Rule['INTERVAL'] } else { 1 }
    if ($interval -lt 1) { $interval = 1 }

    switch ($Rule['FREQ']) {
        'DAILY' {
            if ((($day - $s).Days % $interval) -ne 0) { return $false }
            if ($Rule['BYDAY']) {
                $ok = $false
                foreach ($tok in ($Rule['BYDAY'] -split ',')) {
                    if ($tok -match '(SU|MO|TU|WE|TH|FR|SA)$' -and $script:DOW[$matches[1]] -eq [int]$day.DayOfWeek) { $ok = $true }
                }
                if (-not $ok) { return $false }
            }
            return $true
        }
        'WEEKLY' {
            $wkst = if ($Rule['WKST'] -and $script:DOW.ContainsKey($Rule['WKST'])) { $script:DOW[$Rule['WKST']] } else { 1 }
            $weeks = ((Get-WeekStart $day $wkst) - (Get-WeekStart $s $wkst)).Days / 7
            if (($weeks % $interval) -ne 0) { return $false }
            if ($Rule['BYDAY']) {
                foreach ($tok in ($Rule['BYDAY'] -split ',')) {
                    if ($tok -match '(SU|MO|TU|WE|TH|FR|SA)$' -and $script:DOW[$matches[1]] -eq [int]$day.DayOfWeek) { return $true }
                }
                return $false
            }
            return ($day.DayOfWeek -eq $s.DayOfWeek)
        }
        'MONTHLY' {
            $months = ($day.Year - $s.Year) * 12 + ($day.Month - $s.Month)
            if ($months -lt 0 -or ($months % $interval) -ne 0) { return $false }
            return ((Get-MonthDays $day $Rule $DtStart) -contains $day.Day)
        }
        'YEARLY' {
            $years = $day.Year - $s.Year
            if ($years -lt 0 -or ($years % $interval) -ne 0) { return $false }
            $months = if ($Rule['BYMONTH']) { @($Rule['BYMONTH'] -split ',' | ForEach-Object { [int]$_ }) } else { @($s.Month) }
            if ($months -notcontains $day.Month) { return $false }
            if ($Rule['BYDAY'] -or $Rule['BYMONTHDAY']) { return ((Get-MonthDays $day $Rule $DtStart) -contains $day.Day) }
            return ($day.Day -eq $s.Day)
        }
    }
    $false
}

function Get-IcsOccurrenceIndex($Rule, [datetime]$DtStart, [datetime]$Day) {
    # Only needed for COUNT: how many occurrences happened up to and including $Day.
    $n = 0
    $d = $DtStart.Date
    $end = $Day.Date
    $guard = 0
    while ($d -le $end -and $guard -lt 20000) {
        if (Test-IcsRuleDay $Rule $DtStart $d) { $n++ }
        $d = $d.AddDays(1); $guard++
    }
    $n
}

# --------------------------------------------------------------------------
# the day
# --------------------------------------------------------------------------
function Get-IcsEvents {
<#
    .SYNOPSIS Everything on the calendar that touches $Date, in starting order.
    .OUTPUTS  Start, End, Subject, Location, AllDay, Busy, Oof, Uid
#>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [datetime]$Date = (Get-Date)
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    $dayStart = $Date.Date
    $dayEnd   = $dayStart.AddDays(1)
    $raw      = @(ConvertFrom-IcsText $Text)

    # an occurrence Outlook moved out of its series comes as its own VEVENT with
    # the same UID and a RECURRENCE-ID pointing at the slot it replaces
    $overrides = @{}
    foreach ($e in $raw) {
        if (-not $e.ContainsKey('RECURRENCE-ID')) { continue }
        $uid = Get-IcsProperty $e 'UID'
        $rid = ConvertFrom-IcsDate $e['RECURRENCE-ID'].Value $e['RECURRENCE-ID'].Params
        if ($uid -and $rid) {
            if (-not $overrides.ContainsKey($uid)) { $overrides[$uid] = @{} }
            $overrides[$uid][$rid.At.ToString('yyyyMMddHHmm')] = $true
        }
    }

    $out = @()
    foreach ($e in $raw) {
        if (-not $e.ContainsKey('DTSTART')) { continue }
        if ((Get-IcsProperty $e 'STATUS') -eq 'CANCELLED') { continue }

        $st = ConvertFrom-IcsDate $e['DTSTART'].Value $e['DTSTART'].Params
        if (-not $st) { continue }
        $dtStart = $st.At
        $allDay  = $st.AllDay

        # how long it lasts
        $dur = $null
        if ($e.ContainsKey('DTEND')) {
            $en = ConvertFrom-IcsDate $e['DTEND'].Value $e['DTEND'].Params
            if ($en) { $dur = $en.At - $dtStart }
        } elseif ($e.ContainsKey('DURATION')) {
            $dur = ConvertFrom-IcsDuration (Get-IcsProperty $e 'DURATION')
        }
        if ($null -eq $dur -or $dur.Ticks -lt 0) { $dur = if ($allDay) { [timespan]::FromDays(1) } else { [timespan]::Zero } }

        $subject  = Get-IcsProperty $e 'SUMMARY'
        if (-not $subject) { $subject = '(no subject)' }
        $location = Get-IcsProperty $e 'LOCATION'
        $busy     = (Get-IcsProperty $e 'TRANSP') -ne 'TRANSPARENT'
        $uid      = Get-IcsProperty $e 'UID'

        # Outlook says "out of office" here and nowhere else: such an entry is
        # TRANSP:OPAQUE like any meeting, so TRANSP alone cannot tell them apart
        $oof      = (Get-IcsProperty $e 'X-MICROSOFT-CDO-BUSYSTATUS') -eq 'OOF'

        $starts = @()
        if ($e.ContainsKey('RRULE')) {
            $rule    = ConvertFrom-IcsRule (Get-IcsProperty $e 'RRULE')
            $until   = $null
            if ($rule['UNTIL']) {
                $u = ConvertFrom-IcsDate $rule['UNTIL'] @{}
                if ($u) { $until = $u.At }
            }
            $count = if ($rule['COUNT']) { [int]$rule['COUNT'] } else { 0 }
            $ex    = @{}
            foreach ($x in (Get-IcsDateList $e['EXDATE'])) { $ex[$x.ToString('yyyyMMddHHmm')] = $true }

            # an occurrence may have started on an earlier day and still be running
            $back = [math]::Min(35, [math]::Max(1, [int][math]::Ceiling($dur.TotalDays) + 1))
            for ($i = $back; $i -ge 0; $i--) {
                $cand = $dayStart.AddDays(-$i)
                if (-not (Test-IcsRuleDay $rule $dtStart $cand)) { continue }
                $occ = $cand.Add($dtStart.TimeOfDay)
                if ($until -and $occ -gt $until) { continue }
                if ($count -gt 0 -and (Get-IcsOccurrenceIndex $rule $dtStart $cand) -gt $count) { continue }
                if ($ex[$occ.ToString('yyyyMMddHHmm')]) { continue }
                if ($uid -and $overrides.ContainsKey($uid) -and $overrides[$uid][$occ.ToString('yyyyMMddHHmm')]) { continue }
                $starts += $occ
            }
        } else {
            $starts += $dtStart
        }
        foreach ($r in (Get-IcsDateList $e['RDATE'])) { $starts += $r }

        foreach ($s in ($starts | Sort-Object -Unique)) {
            $end = $s.Add($dur)
            # keep it if it overlaps the day at all; a zero-length event counts
            # when it starts inside the day
            $touches = if ($dur.Ticks -eq 0) { $s -ge $dayStart -and $s -lt $dayEnd }
                       else { $s -lt $dayEnd -and $end -gt $dayStart }
            if (-not $touches) { continue }
            $out += [pscustomobject]@{
                Start    = $s
                End      = $end
                Subject  = $subject
                Location = $location
                AllDay   = $allDay
                Busy     = $busy
                Oof      = $oof
                Uid      = $uid
            }
        }
    }

    # Outlook repeats the same meeting in some feeds; collapse exact duplicates
    $seen = @{}
    $uniq = @()
    foreach ($o in ($out | Sort-Object @{e={-not $_.AllDay}}, Start, Subject)) {
        $k = '{0}|{1}|{2}' -f $o.Start.ToString('yyyyMMddHHmm'), $o.End.ToString('yyyyMMddHHmm'), $o.Subject
        if ($seen[$k]) { continue }
        $seen[$k] = $true
        $uniq += $o
    }
    $uniq
}

Export-ModuleMember -Function Get-IcsEvents, ConvertFrom-IcsText, ConvertFrom-IcsDate,
    ConvertFrom-IcsDuration, ConvertFrom-IcsRule, Test-IcsRuleDay, Expand-IcsFolding
