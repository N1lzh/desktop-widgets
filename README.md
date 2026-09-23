<img width="643" height="176" alt="logo-lockup-dark" src="https://github.com/user-attachments/assets/87520e10-62c1-45f8-8190-a73985b7b4cd" />

# Desktop Widgets

Tiny translucent cards that float on top of everything, have no window frame and
no taskbar entry. You drag them where you like and forget them. They all look
the same, they are all the same width, and they all remember where you put them
(a second or so after you let go, so a Windows restart cannot swallow it).

Four of them so far:

### WorkClock - how long you have worked today

```
  WORKING                 || O X
  2:34                     5:26 left
  =========-------------------------
  of 8:00                 done 12:46
```

* **Worked today** in big digits, **time left** next to it, **progress bar**,
  and the time you are done at the bottom right.
* Counts from the moment it starts and only stops when you **press pause** -
  nothing pauses itself behind your back.
* Everything is remembered - close it, reopen it, the day continues. At midnight
  it starts over.
* Time while the machine sleeps is not counted as work.

### Agenda - today's meetings from your Outlook calendar

```
  NEXT UP                  O @ X
  0:24   Sprint Review
         10:00 - 11:00
  ---------------------------------
  all    Company offsite
  09:00  Daily stand-up
  now    Sprint Review
  13:30  1:1 with lead
  15:00  Design sync
  5 meetings              3:15 booked
```

* A **countdown** to the next meeting, or to the end of the one you are in.
* The rest of the day underneath: what is done is dimmed, what is running says
  `now`, all-day entries say `all`. Hover a line for the full subject, the exact
  time and the location.
* **Always the same size.** More meetings than fit turn into `+3 more later`
  rather than growing the card.
* **Out of office is a blocker, not a meeting.** A daily "Out of Office" from
  16:00 to 23:30 would otherwise own the countdown all afternoon and book 7:30h
  before the day has started, so anything Outlook marks as *out of office*
  disappears: no countdown, no line, not counted as booked. All-day leave and
  holidays stay - they are worth seeing and never counted as booked anyway.
* Refreshes every 10 minutes in the background, keeps the last download so it
  still shows the day when the network is gone, and says `sync failed` when the
  feed did not answer.
* Recurring meetings are expanded properly - weekly series, "every second
  Tuesday", "last Friday of the month", single occurrences you moved or
  deleted, and time zones.

#### Getting the calendar URL

In **Outlook on the web**: Settings -> Calendar -> Shared calendars ->
*Publish a calendar*, pick your calendar, choose **"Can view all details"**,
press Publish, then copy the **ICS** link (not the HTML one).

Right-click the widget -> **Calendar URL...** and paste it. It asks you for it
by itself the first time you start it. Any other ICS feed works too, including
a plain path to an `.ics` file on disk.

### Todos - a checklist you type straight into

```
  TODOS                     O X
  + add a todo...
  ---------------------------------
  [x] buy milk
  [ ] write the report
  [ ] book the flights
  2 open                   1 done
```

* **Type on the top line, press Enter.** That is the whole way in - no dialog,
  no "new todo" button. Escape clears the line again.
* **Click the box** to cross a todo out, click it again if you were too quick.
* Hovering a line shows the **x** that removes it, and when you added it.
* Finished todos stay where they are, crossed out, until you clear them - the
  bin in the corner, or *Clear finished* in the menu. *Show finished* hides them
  instead; the footer keeps counting them either way.
* The card is **as tall as the list** until it reaches nine lines. After that
  the list scrolls and the card stays where it is.
* Every change is written straight to disk, so quitting loses nothing.

### Pet - a small creature that reacts to things

```
  PEACEFUL                   * X
        ( ^ . ^ )
      watching the cursor
  Pixel - day 10         <3<3<3
```

* It has **no needs and it cannot die**. There is nothing to feed, nothing that
  runs out, nothing that blames you for not looking at it for two days. It only
  reacts.
* It reacts to **a meeting about to start** (`( O _ O )  Sprint Review in 4 min`,
  read from the Agenda's cached calendar), to **being in one** (`shh...`) and to
  **it ending**, to **a todo you ticked off** (`one down, 3 to go`) or **added**,
  to **the WorkClock being paused**, to **reaching your daily target** and to
  **working past it**, to **lunch** from 11:45, to **the machine coming back**
  after it was suspended, to **midnight passing**, and on a laptop to **being
  unplugged** and to **the battery running low**.
* **It also does things on its own.** Every few minutes, when nothing else is
  going on, it stretches, sneezes, chases its tail, sings, sniffs the taskbar,
  finds a crumb or takes a micro nap - and jumps on the spot for the good ones.
  Everything it says comes out of a pool, so it rarely says the same thing
  twice, and it blinks at random - sometimes twice in a row.
* **It sleeps outside the day**: from **16:00 until 6:00**, and any time the
  keyboard and mouse have been untouched for five minutes. `z z z`, and it wakes
  up when you come back.
* **Click the face to pet it**; anywhere else on the card still drags. Pet it
  too fast and it goes `( @ _ @ )`. Right-click for a treat, or to give it a
  name.
* **Three creatures to pick from** under *Style*, switched live. They all react
  to exactly the same things and say the same things - only the face differs:

```
  Classic   ( ^ . ^ )   ( o _ o )   \ ( ^ o ^ ) /   ( @ _ @ )   ( u _ u )
  Cat       ( ^ w ^ )~  ( = w = )~  \ ( ^ w ^ ) /   ( @ w @ )   ( u w u )~
  Robot     [ o _ o ]   [ = _ = ]   \ [ ^ _ ^ ] /   [ x _ x ]   [ u _ u ]
            peaceful    in a mtg    happy           dizzy       tired
```
* The **first start is an egg** that wobbles for forty seconds and then hatches.
  The footer counts the days since; the hearts fill up as you pet it.
* It never pops up, never steals focus and never asks for anything. Between
  reactions it blinks, breathes and now and then mutters to itself.

## Run it

Windows, no installation, nothing to build:

```
install.cmd              double-click: copies everything to %LOCALAPPDATA%
                         and starts all of them with Windows
```

or just run them where they are:

```
wscript Widget.vbs workclock                                   Windows
wscript Widget.vbs agenda
wscript Widget.vbs todos
wscript Widget.vbs pet
./widget agenda                                                from WSL
powershell -NoProfile -ExecutionPolicy Bypass -File widgets\agenda.ps1
```

`install.ps1 -Uninstall` removes it again (your recorded times and settings are
kept). `install.ps1 -Widgets agenda` installs only one of them.

Starting them straight out of a WSL folder works - that is what `./widget` does -
but the Startup shortcut then points at `\\wsl.localhost\...`, which only
resolves once the distro is running. Install them if you want them up reliably
at logon.

## Using them

| | |
|---|---|
| drag anywhere on the card | move the widget |
| hover the card | the buttons appear |
| scroll on the card | more / less transparent |
| right-click | menu |

Every widget has the same bottom half of its menu:

* **Background > Lighter / Darker** - six steps, from nearly invisible to an
  almost solid card. **Scrolling** on the widget does the same thing.
* **Background > Blur what is behind** - frosts the card instead of letting the
  desktop through. Prettier on busy backgrounds, but it stops being see-through.
  Off by default, toggles live.
* **Background > Windows acrylic** - the Windows 11 system material. Windows
  picks the tint there, so it looks close to a solid card. Off by default.
* **Start with Windows** - puts a shortcut in the Startup folder that points at
  the folder the widget is running from. Move or rename that folder and the
  shortcut goes dead; the next time the widget runs it notices and repoints the
  shortcut at itself. Running `install.cmd` is the sturdier option: it copies
  the widgets into `%LOCALAPPDATA%` first, so the shortcuts do not depend on
  where the sources happen to live (or, from WSL, on the distro being up).
* **Quit**

On top of that, WorkClock has **Pause/Resume**, **Adjust worked time...** (if
you started working before you started the widget - accepts `3:20`, `200m` or
`3.5`), **Set daily target...** and **Reset today**; Agenda has **Refresh now**,
**Calendar URL...**, **Refresh every...** and **Show all-day events**; Todos has
**Show finished**, **Clear finished** and **Remove all todos...**; Pet has
**Give a treat**, **Name it...**, **Style** and **Start over as an egg...**.

## Where things live

```
lib/Widget.psm1        the card itself: the window, the tints, blur, dragging,
                       the shared menu entries, the state file, autostart
lib/Ics.psm1           just enough iCalendar to answer "what is on today?"
widgets/workclock.ps1  the work clock
widgets/agenda.ps1     the meeting list
widgets/todos.ps1      the checklist
widgets/pet.ps1        the creature
widgets.json           which widgets exist (the installer reads this)
Widget.vbs             starts one of them without a console window
widget                 the same, from a WSL shell
install.ps1 / .cmd     copy to %LOCALAPPDATA%\WorkClock\app + shortcuts
```

Settings and recorded times live in `%LOCALAPPDATA%\WorkClock\`: `state.json`
for the clock, `agenda.json` plus a cached `agenda.ics` for the meetings,
`todos.json` for the list and `pet.json` for the creature. The Pet reads the
other three to know what to react to.

After every refresh the Agenda also writes `agenda-today.json`: today's
meetings exactly as the card shows them (recurrences expanded, out of office
dropped), with ISO start and end times. It is meant for other tools, like a
terminal greeting in WSL, so they do not have to parse the ICS themselves.

Set `WORKCLOCK_DEBUG=1` to get a log at `%TEMP%\workclock.log`.

## Adding another widget

Copy `widgets/workclock.ps1`, keep the top and bottom, replace the middle:

```powershell
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'lib\Widget.psm1') -Force

$script:S = Initialize-Widget -Id 'mything' -Name 'My Thing' -ScriptPath $PSCommandPath `
    -Defaults ([ordered]@{ something = 1 })

$win = New-WidgetWindow -Body $xamlRowsGoHere -MenuItems $myMenuItemsXaml   # -Height for a fixed card
$script:UI = Get-WidgetElements $win @('MyTextBlock')
Register-WidgetChrome $win
[void]$win.ShowDialog()
```

Add it to `widgets.json` and the installer picks it up. `New-WidgetWindow` gives
you the same 258px card, the `Glyph` button style, the context menu and the
`Card` / `Controls` element names; `Get-WidgetPalette` gives you the colours the
other widgets paint with.
