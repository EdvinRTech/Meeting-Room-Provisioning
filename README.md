# Meeting Room Provisioning Tool

A GUI wizard for provisioning and editing Microsoft Teams meeting room
accounts across any Exchange Online / Microsoft Graph tenant. Replaces
`ExchangeOnline/New-MTRAutomatedGUI.ps1`'s hardcoded, single-tenant script
with a dynamic one: pick or create a Room List, pick or create the
Conditional Access exclusion group, optionally set up SSPR exclusion, type
in room names, set place info, choose calendar processing rules in plain
language, and create (or edit) everything in one run.

## Requirements

- Windows with PowerShell 5.1 available (the tool relaunches itself under
  PowerShell 7 if installed - see [Troubleshooting](#troubleshooting)). If
  PowerShell 7 isn't installed, the first launch offers to install it
  automatically via winget (one confirmation prompt, then it just works
  from there on) - see [Automatic PowerShell 7 install](#automatic-powershell-7-install).
  Say no, or if winget isn't available, and the tool still runs, but Graph
  sign-in is likely to fail (see below).
- Local admin rights - the tool self-elevates (UAC prompt) on every launch.
- An Entra account with:
  - **Exchange Administrator**, and
  - **Password Administrator**, **User Administrator**, or **Global
    Administrator**.

  Exchange Administrator alone gets you through most of the wizard, but
  the password-setting step fails with `Authorization_RequestDenied` -
  setting another user's password via Graph requires a privileged
  directory role in addition to the API permission grant; the role and
  the permission are checked separately by Microsoft's design.

  If you hold Global Administrator but still get `Authorization_RequestDenied`,
  check the Connect step's success message: it lists the signed-in
  account's **currently active** directory roles, queried live. In
  tenants using PIM, a role that's *eligible* but not *activated* for the
  current sign-in doesn't count, even though the portal still shows you
  holding it - activate it for the session and reconnect.

## Getting started

Double-click `M365 Meeting Room Tool.exe`. That's the whole
"getting started" step - no right-click menu, no execution policy to
think about for this file specifically.

Or, from a PowerShell prompt:

```powershell
.\App\Start-MeetingRoomProvisioning.ps1
```

(`App` is a hidden folder - see [Architecture](#architecture) - so toggle
Explorer's "Show hidden items", or just type the path, to reach it.
"Run with PowerShell" from its right-click menu also works.) Either way,
expect a UAC prompt on every launch - the tool relaunches itself
elevated, in STA mode, and under PowerShell 7 if available, before
anything else runs. See [The .exe launcher](#the-exe-launcher) for what
that file actually is and why it exists alongside the `.ps1`.

The first Connect click installs required modules (Exchange Online
Management, Microsoft.Graph submodules) and signs you in to both Exchange
Online and Microsoft Graph via the normal Windows sign-in popup (WAM) -
there's no username/password field, since neither service supports plain
password auth for this kind of sign-in anymore. A progress bar tracks the
whole Connect flow at the granularity of individual operations, not
coarse phases: disconnecting Exchange Online and Graph separately (2),
removing each old version of each required module actually found on the
machine (varies - could be zero), installing each of the 7 required
modules (7), importing each of them (7), then signing in and loading
initial tenant data (2 more) - so the bar's fill genuinely tracks
remaining work instead of jumping in a few big, uneven steps, and the
status text names the specific module currently being handled ("Removing
old modules... (Microsoft.Graph.Users)"). Module install/removal failures
are still recorded in full detail and surfaced if the step actually
fails, but per-module success/failure lines (which can include a raw,
sometimes multi-sentence .NET exception message) are no longer flashed
past one at a time on that single status line.

### Distributing to colleagues

Copy the whole `MeetingRoomProvisioning` folder - `M365 Meeting Room Tool.exe`
plus the `App` folder next to it (hidden by default; still there, still
needs copying) - nothing else to install. The `.exe` needs `App\Start-MeetingRoomProvisioning.ps1`
sitting right next to it (see [The .exe launcher](#the-exe-launcher)), so
don't hand out the `.exe` on its own.

**A ZIP downloaded via a browser (e.g. GitHub's "Download ZIP") carries
Windows' "Mark of the Web", and every file extracted from it inherits
that flag too** - not just the main `.ps1`. Under the default execution
policy this blocks the script from running *at all*, with no visible
error: PowerShell refuses to load a blocked script before its first line
ever executes, so nothing appears - not even a console flash, and not
even the startup-error log this tool otherwise writes on failure, since
the script itself never starts.

**Unblock the ZIP itself before extracting it** - confirmed as the fix
that actually works, unlike unblocking the already-extracted files
afterward, which did not reliably clear the block in practice:

- Right-click the `.zip` → Properties → check **Unblock** at the bottom →
  OK, or
- `Unblock-File .\Meeting-Room-Provisioning-main.zip` (adjust the
  filename to whatever you downloaded)

Then extract as normal - everything that comes out of an unblocked ZIP is
unblocked too. If you've already extracted a blocked ZIP,
`Get-ChildItem -Path . -Recurse | Unblock-File` on the extracted folder
*should* be equivalent, but re-downloading and unblocking the ZIP first
is the version that's actually been confirmed to work. The `.exe` is
subject to the exact same Mark-of-the-Web blocking as every other file in
the folder - unblocking the ZIP first covers it too.

## The .exe launcher

`M365 Meeting Room Tool.exe` exists purely so the tool can be
double-clicked directly, instead of needing "right-click > Run with
PowerShell" on the `.ps1`. It is **not** a compiled copy of the actual
application - it's a ~20-line stub (`App\Launcher.ps1`, compiled via
[PS2EXE](https://github.com/MScholtes/PS2EXE)) whose only job is to find
`App\Start-MeetingRoomProvisioning.ps1` - in the hidden `App` folder next
to it, see [Architecture](#architecture) - and hand off to a real
`pwsh.exe` (or `powershell.exe`, if PowerShell 7 isn't installed) process
running it.

That indirection exists for a concrete, tested reason: a PS2EXE-compiled
executable always hosts Windows PowerShell 5.1 Desktop internally,
regardless of which PowerShell version builds it - confirmed empirically
while building this (`$PSVersionTable.PSEdition` reads `Desktop` inside
the compiled `.exe` even when it's built from `pwsh.exe`, and there is no
PS2EXE option to change that). That's exactly the engine this tool's own
relaunch logic exists to get away from, for the Graph SDK's
`GetTokenAsync ... lacks an implementation` bug (see
[Troubleshooting](#troubleshooting)). Compiling the real ~1300-line
application would have permanently baked that limitation in. Handing off
to a real `pwsh.exe -File` process instead means every one of the real
script's own mechanisms - elevation, STA, PowerShell 7 preference/auto-install,
startup error logging - applies completely unchanged, exactly as if you'd
run the `.ps1` yourself; the launcher doesn't duplicate or reimplement any
of it.

Also confirmed empirically, not assumed: `$PSScriptRoot` and
`$PSCommandPath` are both **empty** inside a compiled `.exe` (so the
launcher locates its own folder via
`[System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName`
instead), and PS2EXE's `-noConsole` build option hangs indefinitely here
rather than exiting once the launcher's work is done - confirmed by
actually running the built `.exe` repeatedly, not just reading PS2EXE's
own documentation. `Build-Exe.ps1` (see below) deliberately does not use
it; the cost is a brief, harmless console flash while the launcher runs.

The launcher also elevates and hands off in **one** `Start-Process` call,
already carrying `-RelaunchedForGui` and `-WindowStyle Hidden` (plus
`-Verb RunAs` only when not already elevated) - i.e. it jumps straight to
the exact end state `Start-MeetingRoomProvisioning.ps1`'s own relaunch
logic targets, so that script finds nothing left to do and takes no
further hop. An earlier version instead launched a plain, non-elevated hop
and let the real script elevate *itself* a second time from there - which
left a visible, elevated console full of module-install output behind
(worked its way to a WAM sign-in popup with a full console window sitting
behind it). Only a *single* `-Verb RunAs` combined with `-WindowStyle
Hidden`, issued directly from the process actually being double-clicked,
has been confirmed to reliably hide the console - the same nested
combination one relaunch hop deeper did not.

**Rebuilding it:** the `.exe` is not generated automatically - it has to
be rebuilt and re-committed by hand after any change to `Launcher.ps1`
(changes to `Start-MeetingRoomProvisioning.ps1` itself do *not* need a
rebuild, since the launcher hands off to that file by path rather than
embedding it). Run it from the repo root - it writes the `.exe` one level
up from `App`, so it lands next to `App` rather than inside it:

```powershell
.\App\Build-Exe.ps1
```

Installs the `ps2exe` module (CurrentUser scope) if it isn't already
present. The very first time a freshly-built, unsigned `.exe` runs on a
given machine, Windows Defender/SmartScreen commonly scans it before
letting it start, which can look like a 5-15 second hang on that first
launch specifically - confirmed while building this; every run after that
first one starts and hands off promptly. Not a bug, just worth knowing
before assuming something's wrong. A real code-signing certificate would
remove that first-run delay (and the more prominent "Windows protected
your PC" SmartScreen prompt an unsigned `.exe` downloaded from the
internet is also likely to show) but is a separate, ongoing paid process,
not something this tool sets up for you.

### The icon

`AppIcon.ico` is generated by `New-AppIcon.ps1` - pure
[System.Drawing](https://learn.microsoft.com/dotnet/api/system.drawing)
(GDI+), no external image tool or downloaded asset: a simple meeting-room
display screen (with a small camera dot) on a stand, in white and the
brand's accent blue, on a rounded navy square. It hand-assembles a proper multi-resolution
`.ico` (PNG-compressed entries per size - the format Windows has
supported since Vista - at 16/32/48/64/256px) rather than relying on
`Bitmap.Save(..., ImageFormat.Icon)`, which only writes a single
resolution.

It's used in two places, wired independently:

- **`M365 Meeting Room Tool.exe`** embeds it via `Build-Exe.ps1`'s
  `-iconFile` - this needs the `.exe` rebuilt (`.\App\Build-Exe.ps1`) to
  pick up any change to `AppIcon.ico`.
- **The running window** (title bar / taskbar / Alt-Tab) sets it from
  code in `Start-MeetingRoomProvisioning.ps1`, right after loading the
  XAML, rather than via XAML's own `Icon="..."` attribute - confirmed
  during testing that a relative path there would be ambiguous, since
  this XAML loads from a loose file via `XamlReader.Load` (no compiled
  `pack://` resource resolution) and would resolve against whatever the
  process's current working directory happens to be, not necessarily
  `UI\MainWindow.xaml`'s own folder. An absolute path built from
  `$ScriptRoot` (the same variable every other file this tool loads
  already uses) sidesteps that. This one just needs the tool relaunched,
  no rebuild.

Regenerate with `.\App\New-AppIcon.ps1` after changing the design in that
file, then re-run `.\App\Build-Exe.ps1` for the `.exe`'s copy specifically.

### The sidebar logo

The real Asurgent brand mark (the winged "A", per the Identity leveranspaket
spec) sits in the sidebar next to the "Asurgent" wordmark, where the design
originally had a plain placeholder square. `App\UI\AsurgentLogo.png` is the
icon cropped out of the full lock-up down to just the mark (transparent
background, no baked-in wordmark text - the existing `TextBlock` next to it
already renders "Asurgent" in the brand's own serif); `App\UI\AsurgentLogo-full.png`
keeps the original full image (mark + wordmark) around in case a different
crop is ever needed later. Same "set from code, not XAML" reasoning and
mechanism as the window icon above - `imgLogo` in `MainWindow.xaml` is a
plain, source-less `<Image>` that `Start-MeetingRoomProvisioning.ps1` points
at `AsurgentLogo.png` right after loading the XAML.

## The wizard, step by step

1. **Connect** - signs in to Exchange Online and Microsoft Graph (scopes:
   `User.ReadWrite.All`, `Group.ReadWrite.All`, `Policy.Read.All`,
   `Policy.ReadWrite.ConditionalAccess`, `Directory.Read.All`,
   `Organization.Read.All`). Shows a license summary for existing room
   mailboxes (informational only - this tool never assigns licenses).
2. **Mode** - **Create new rooms** or **Edit existing rooms** (change
   calendar processing, place info, Room List, and/or password on rooms
   that already exist). See [Edit mode](#edit-mode) below.
3. **Room List** - pick an existing one or create a new one (its address
   is generated automatically from the tenant's default domain). In Edit
   mode, membership can also be left unchanged.
4. **Conditional Access exclusion group** *(both modes)* - pick an
   existing group excluded from at least one CA policy, or create a new
   one. Either way it's synced against **every CA policy that currently
   exists**, not just the ones that existed when the group was created -
   see [CA exclusion group sync](#ca-exclusion-group-sync). Edit mode gets
   the same choice as Create mode, since a room being edited may never
   have been added to the group in the first place (it might predate this
   tool managing that, or the group might not have existed yet) -
   `Add-RoomToGroup` is idempotent, so a room already in the group is left
   alone.
5. **SSPR exclusion** *(both modes)* - shown only if Self-Service Password
   Reset is enabled tenant-wide. Edit mode gets the same "create the group
   if missing" option as Create mode - editing an existing room is exactly
   how one created before this feature existed gets excluded. See
   [SSPR exclusion](#sspr-exclusion).
6. **Rooms** - Create mode: type room names and pick a domain. Edit mode:
   pick one or more existing room mailboxes from a list.
7. **Place info** - maps to `Set-Place`. A blank field is left out of the
   command entirely (not sent as empty), so in Edit mode a blank field
   means "leave as-is". Country is a dropdown of proper names mapped to
   the 2-letter ISO codes `Set-Place -CountryOrRegion` actually expects,
   built from .NET's `System.Globalization.RegionInfo`.
8. **Calendar processing** - Standard mode uses the tool's built-in
   defaults; Custom mode asks plain-language questions and translates
   them to `Set-CalendarProcessing` parameters (see [Calendar processing](#calendar-processing)).
   Edit mode adds a "don't change" option, selected by default.
9. **Review & create/apply** - shows a summary, a password field (see
   [Room password](#room-password)), then runs the operation with a
   colored log and retry progress. See [Retries and cancelling](#retries-and-cancelling-a-run).

Several steps (Room List, CA exclusion group, SSPR exclusion, Place info,
Calendar processing) have a collapsed-by-default "what does this actually
do?" expander with a more technical explanation of the underlying
Graph/Exchange mechanics - click it to expand. The existing-license
summary on the Connect step is colored the same info-cyan as the run log's
informational lines, for the same reason: it's context, not a
warning or a result.

## CA exclusion group sync

`New-ConditionalAccessExclusionGroup` and picking an existing group both
call `Sync-GroupExclusionAcrossConditionalAccessPolicies`
(`RoomProvisioning.Graph.psm1`), which excludes the group from every CA
policy that exists **right now** - not just the ones present when the
group was created. Without this, a CA policy added later by another admin
would silently start applying to meeting room accounts. It runs on every
provisioning run, so the exclusion list self-heals over time instead of
needing a manual audit; the log reports how many policies were already
excluded vs. newly excluded.

Each per-policy update retries a few times a couple of seconds apart: a
policy that was itself just created can briefly 404 on write (directory
replication lag), and Graph load-balances reads across replicas that
converge independently, so a policy visible in the initial listing can
still occasionally reject an update moments later.

The "existing excluded group" list on the CA step (and the SSPR step's
found/missing state) is queried once at Connect and then re-queried again
after every Create/Apply run finishes - not on every step navigation, to
avoid a live Graph call on every Back/Next click. Without that refresh, a
group created during one run wouldn't show up as "existing" if the admin
clicked Back afterward to process another batch of rooms in the same
session, instead of closing and reopening the tool.

## SSPR exclusion

Meeting room accounts have no owner and shouldn't be reachable through
Self-Service Password Reset - mainly because SSPR normally requires a
user to have already registered MFA methods to verify their identity
during a reset, and a room mailbox has no human owner who could ever
register or complete an MFA challenge. Microsoft Graph exposes only one
SSPR
setting via API: a tenant-wide on/off boolean (`AllowedToUseSSPR` on
`policies/authorizationPolicy`, read via `Test-SelfServicePasswordResetEnabled`).
**There is no API to read or set SSPR's group scope** (All users vs.
specific security groups, or which ones) - confirmed against the stable
and beta Graph SDKs and raw REST calls. That's a Microsoft platform gap,
not a limitation of this tool, and it caps what this feature can actually
do: whichever option you pick, you still need to confirm the real scope
yourself in the Entra admin center (**Password reset > Properties**) -
the wizard and the final result both say so. The step's status line spells
this out explicitly whenever SSPR is on ("...but Graph can't tell whether
it's scoped to All users or to Selected groups, only that it's on"),
rather than just saying "enabled" and leaving that gap implicit.

When `AllowedToUseSSPR` is `false` - SSPR off tenant-wide, nothing to
exclude rooms from - the step just says so and the "create the group"
checkbox isn't offered at all; there's nothing it would accomplish.

What the tool does, within that limit, is built around a single
standard-named group, **"SSPR Users"** (`Get-SsprGroupDisplayName` in
`RoomProvisioning.Graph.psm1`) - fixed rather than admin-typed, since the
tool needs to find it by name on every run without asking again:

- **Looks up "SSPR Users" by name on every run** (`Get-SsprExclusionGroup`),
  Create or Edit, whenever SSPR is enabled. If it isn't found, that run's
  SSPR step is skipped entirely - no error, just a log line saying so.
- **Offers to create it** (via a checkbox on the SSPR step, in either
  Create or Edit mode) if it doesn't already exist, with this membership
  rule:

  ```
  (user.assignedPlans -any (assignedPlan.servicePlanId -ne "" -and assignedPlan.capabilityStatus -eq "Enabled"))
  and (user.userType -eq "Member")
  and (user.accountEnabled -eq true)
  ```

  i.e. licensed, active, Member-type (not guest) accounts - no room
  exclusions yet at creation time. Because Graph can't tell whether SSPR
  is already scoped to some other, differently-named group, the wizard
  asks you to check Entra admin center > Password reset > Properties
  yourself before checking this box - creating a new group here does
  nothing for SSPR if a different group is already targeted.
- **Excludes each room by UPN** as it's created or edited, in both Create
  and Edit mode (Edit mode isn't gated on the password-reset checkbox, so
  it also backfills rooms edited before this feature existed): appends
  `and (user.userPrincipalName -ne "room@domain.com")` to the group's
  rule, one clause per room, skipping any room whose clause is already
  present. Existing rule logic is never touched - only appended to.

Because every room in a run updates the same group sequentially,
`Start-MeetingRoomProvisioning.ps1` tracks the rule's current value
**locally** across rooms instead of re-reading it from Graph between
updates - a fresh read could land on a replica that hasn't caught up with
the previous room's write yet (see the replication-lag notes throughout
this doc) and silently clobber it. `Get-SsprExclusionGroup` still retries
its lookup (`-ConsistencyLevel eventual` plus a few attempts), and
`Add-RoomToSsprExclusionRule`'s single write is wrapped in the same
`Invoke-WithRetryProgress` used for every other per-room Graph call.

A rule that accumulates one clause per room indefinitely will eventually
approach Entra's dynamic membership rule length limit in a tenant with a
very large number of rooms - not a concern at normal scale, but worth
knowing if this group has been in use for years across hundreds of rooms.

## Room password

There is no default or hardcoded password. The Review step has a masked
password field plus a confirmation field, required before every Create
run and before every Edit run with "Reset password for these rooms"
checked. Every room processed in one run shares that password, but
nothing is saved between runs - closing and reopening the tool prompts
again. (An earlier version defaulted to a fixed password for every
tenant unless an admin remembered to change it - removed as a real
security problem for a tool reused across customer tenants.)

Setting the password is actually two separate Graph calls, each its own
retried step with its own run-log line: disabling password expiration on
the account first (`Set-RoomPasswordPolicy`, logged as "Password policy
set after N attempt(s)." or a FAILED line naming the error), then setting
the password value itself (`Set-RoomPassword`, "Password set/reset after
N attempt(s)."). They used to be one combined call, which meant a
password-policy failure and a password-value failure were
indistinguishable in the log - a room silently left subject to normal
expiration is a meaningfully different, and worse, outcome than the
reverse, so they're now reported separately. A policy-step failure
doesn't skip the password-value step; they're independent, and only
rooms where the password *value* was actually set end up in the
clipboard/result-card summary below.

Once a run finishes, the password and every room address it was
successfully set on (only those - a room whose password attempt failed
doesn't actually have it, so it's left out rather than listed
incorrectly) are copied straight to the clipboard, and shown the same
way in a read-only, selectable text box on the result card, as a
fallback for whenever the clipboard doesn't reach wherever it needs to be
pasted (e.g. a remote session with clipboard redirection off). Nothing is
copied or shown when no password was actually set this run (Edit mode
with the reset checkbox left unchecked, or every reset attempt failed).

### Room name to email address

A room's email/username local part is derived from its display name by
transliterating accented Latin letters to their unaccented base first -
`é`/`è`/`ê` → `e`, `å`/`ä`/`à` → `a`, `ö`/`ô` → `o`, and so on for the
whole range Unicode covers this way (via NFD normalization + stripping
combining diacritic marks, not a hand-typed table of four Nordic
letters) - then stripping anything still left that isn't a letter,
digit, `-`, or `.`. So "Örebro - Café" becomes `orebrocafe`, not
`rebrocaf` (which is what silently deleting every accented character,
the previous behavior, produced instead).

## Edit mode

A second mode for changing settings on rooms that already exist, narrower
in scope than Create mode:

- No mailbox creation - that's the only thing genuinely Create-mode-only.
  Both the Conditional Access group step and the SSPR exclusion step apply
  in Edit mode too, so a room that predates either feature (or was created
  before the relevant group existed) can be brought in line with a
  regular edit run instead of needing to be recreated.
- Room List, calendar processing, and password are all optional, each
  defaulting to "don't change" - nothing you don't explicitly opt into
  gets touched.
- Place info always applies, with the same "blank = leave as-is" rule as
  Create mode.
- Rooms are picked from a live, multi-select list, so one run can apply
  the same change to several rooms at once.

## Calendar processing

Standard mode's defaults match the original script's, with one cleanup:
`AddOrganizerToSubject` always stays `$false` (a Teams Rooms panel has no
use for the organizer's name in the subject either way) rather than being
tied to the "private room" question, since it wasn't actually
privacy-specific in the original script.

Custom mode also exposes settings beyond the original spec: recurring
meetings allowed, attachment stripping, external meeting requests,
removing the "Private" flag on synced meetings, a custom auto-reply
message, and delegate-approval bookings.

## Retries and cancelling a run

Password-policy, password-setting, group membership, place-info, and
SSPR-exclusion calls all retry (10 attempts, 20s apart by default) because
a just-created or
just-changed account isn't always immediately visible to Graph/Exchange
writes. The progress bar's max is the retry cap, so it reflects attempts
remaining rather than spinning generically - and once a call actually
succeeds, the caller explicitly sets the bar to its own maximum so it
visibly fills to 100% rather than being left sitting at whichever attempt
number it happened to succeed on (`Invoke-WithRetryProgress` itself has no
notion of "done", just "attempt N of M" - completing the bar is the
caller's job). The password is reported as set **only for rooms it was
actually confirmed set on** - success/failure is tracked per room, and the
result card turns amber instead of green if anything failed.

A **Cancel** button appears once a run starts, with a confirmation dialog
("rooms already fully processed keep their changes; the current one may
be left partially configured"). Because the tool has no background
thread, a plain `Start-Sleep` during a retry wait would freeze the whole
window - `Invoke-WithRetryProgress` (`RoomProvisioning.Common.psm1`) cuts
each wait into ~200ms chunks and pumps the WPF dispatcher between them, so
a Cancel click is processed within about one chunk instead of only
between whole attempts. Cancelling is checked at the start of each room
and inside each retry wait; a step already talking to Graph/Exchange
always finishes that one network call first.

The run log is colored by outcome, not tracked as separate state: lines
containing "FAILED" or "Cancelled" are red, recognized completions
("created", "configured", "added", "set after", etc.) are green,
everything else is cyan.

The last step's "Next" button becomes "Exit" instead of disappearing, and
asks for confirmation before closing - the same pattern as Cancel, so an
out-of-habit click doesn't close the app by surprise.

## Architecture

```
M365 Meeting Room Tool.exe          Double-click launcher stub - see "The .exe launcher"
README.md                           This file
App/                                Everything else - hidden (Windows Hidden attribute) so the
                                     distributed folder shows only the .exe (and this README)
  Start-MeetingRoomProvisioning.ps1 Entry point: loads XAML, wires events, orchestrates creation
  Launcher.ps1                      Source the .exe above is compiled from
  Build-Exe.ps1                     Rebuilds the .exe from Launcher.ps1 (run by hand, not automatic)
  AppIcon.ico                       App icon - see "The icon"
  New-AppIcon.ps1                   Generates AppIcon.ico (run by hand, not automatic)
  UI/MainWindow.xaml                Window layout only, no logic
  Modules/
    RoomProvisioning.Common.psm1      Generic retry-with-progress helper
    RoomProvisioning.Connections.psm1 Module install + EXO/Graph sign-in
    RoomProvisioning.Exchange.psm1    Room List, mailbox creation, Set-Place
    RoomProvisioning.Graph.psm1       CA policy exclusion group, SSPR exclusion, domains, license check, password
    RoomProvisioning.CalendarLogic.psm1  Plain-language -> Set-CalendarProcessing mapping
```

`App` being hidden is purely cosmetic - a plain Windows folder attribute, not a
security boundary. `Launcher.ps1` re-applies it on every launch since it's
filesystem metadata, not file content, so it doesn't survive a re-zip/re-extract
or a fresh `git clone` on its own. Toggle Explorer's "Show hidden items" (or
`Get-ChildItem -Force`/`dir /a`) to see inside it; nothing about editing or
running the tool requires unhiding it first.

## Troubleshooting

**Nothing happens at all when launching - no window, no error.** Check
first whether the files are blocked (Mark of the Web) - see
[Distributing to colleagues](#distributing-to-colleagues). This is the
most likely cause specifically when the tool was downloaded as a ZIP
(e.g. GitHub's "Download ZIP") rather than `git clone`d or copied from an
already-trusted location: a blocked script is refused by the default
execution policy *before its first line ever runs*, so there's no window,
no console flash, and - importantly - not even an entry in the startup
log described below, since the script itself never starts far enough to
write one. **Unblock the ZIP itself before extracting it** - confirmed in
practice as the fix that actually works, where unblocking the individual
files after extraction did not.

If the files aren't blocked and it's still silent: every launch
relaunches itself into a hidden process (see [Getting started](#getting-started))
to reach a consistent elevated/STA/PowerShell-7 state, so a failure
anywhere between that relaunch and the window actually appearing - a
missing `Modules\` or `UI\` file, a XAML parse error, an assembly that
isn't available on this machine - has nowhere visible to show up by
default. That whole span is wrapped in a handler that writes full details
to `%TEMP%\MeetingRoomProvisioning-startup-error.log` and shows a message
box - check that log file first; it names the exact failure. If even the
message box never appears (i.e. truly nothing at all, not even after
several seconds), the most likely explanation is
`Add-Type -AssemblyName PresentationFramework` itself failing before that
handler can even show a message box - WPF isn't available on every
Windows configuration (Server Core, and PowerShell 7 on ARM64 in
particular, since there's no ARM64 Windows Desktop runtime for it) -
in which case the log file is still the thing to check, since it's
written before the message box is attempted.

**`Authorization_RequestDenied` on the password step, despite holding the
right role and a fresh token.** This app requests six Graph scopes at
once, and the interactive WAM consent flow for a multi-scope request can
fail to fully complete server-side even though `(Get-MgContext).Scopes`
reports every scope as "granted" - that property reflects what was
requested/cached client-side, not the real server-side consent record.
`Connect-RoomProvisioningServices` now calls `Test-RequiredGraphScopesGranted`
right after connecting, which checks `Get-MgOauth2PermissionGrant` (the
actual server-side record) and fails immediately with a specific message
naming the missing permission and where to fix it (Entra admin center →
Enterprise applications → the app → Permissions → **Grant admin
consent**), instead of surfacing as a delayed, cryptic 403.

**`Method GetTokenAsync ... lacks an implementation`.** The Microsoft
Graph SDK ships a separate "Desktop" build of `Azure.Core` for Windows
PowerShell 5.1's .NET Framework runtime, which is incompatible with
recent SDK releases - PowerShell 7's .NET (Core) build doesn't have this
bug. The tool's startup relaunch logic actively prefers `pwsh.exe` over
whatever launched it for exactly this reason, and offers to install it
automatically if it's missing - see
[Automatic PowerShell 7 install](#automatic-powershell-7-install). If
that was declined or failed, the tool falls back to Windows PowerShell
and the Connect step shows a warning recommending
`winget install Microsoft.PowerShell`.

## Automatic PowerShell 7 install

If `pwsh.exe` isn't found once the tool is running elevated, it offers to
install PowerShell 7 for you via winget (Windows Package Manager) before
doing anything else - a single Yes/No prompt ("Install it now via
winget?"), so a machine that's never run this tool before doesn't need a
separate manual install step first. This only happens once elevation is
already confirmed (installing software needs admin rights too, so asking
before that would just fail) and only under Windows PowerShell (a
pwsh-hosted relaunch of this same script would otherwise ask again on
every single launch for no reason).

Saying yes runs `winget install --id Microsoft.PowerShell --source winget
-e --silent --accept-source-agreements --accept-package-agreements`,
shown in a normal (not hidden) window - unlike every other relaunch this
tool does, which are all hidden, since a silent window here with no
visible progress for what can be 10-60+ seconds would just look hung.
Once winget finishes, the tool re-checks for `pwsh.exe` by its default
per-machine install path (`%ProgramFiles%\PowerShell\7\pwsh.exe`) in
addition to `Get-Command` - an installer updates the registry's
`Environment` key, not any already-running process's in-memory
`$env:PATH`, so `Get-Command` alone wouldn't find a copy that was just
installed moments ago by a sibling process. If PowerShell 7 is found
afterward, the tool's normal relaunch-to-pwsh logic picks it up exactly
as if it had been there all along - no separate restart needed by hand.

Declining the prompt, winget not being available, or the install itself
failing are all non-fatal: the tool logs/shows a brief explanation and
continues under Windows PowerShell 5.1, same as if this feature didn't
exist at all.

**`AADSTS500014` ("service principal ... is disabled") during Exchange
sign-in, despite the account being fine.** Recent `ExchangeOnlineManagement`
versions default interactive sign-in to the Windows account broker (WAM),
which can silently sign in with whatever Windows account is already
logged in instead of prompting for the admin account you're typing in.
The tool pins `ExchangeOnlineManagement` to **3.6.0**, which predates that
default and reliably prompts for the intended account.

**Assembly-loading errors after Connect.** The tool removes every
existing install of `ExchangeOnlineManagement` and each `Microsoft.Graph.*`
submodule (both CurrentUser and AllUsers scope) before reinstalling at
fixed, known-good versions (`$Script:GraphModuleVersion` /
`$Script:ExchangeOnlineManagementVersion` in `RoomProvisioning.Connections.psm1`),
then imports each by its exact installed path rather than by name+version
search. If an error persists, another product on the machine (e.g. an
`Az.*` module, which ships its own `Azure.Core`) may be loading its
assemblies first - the Connect step's error display lists every loaded
`Azure.Core`/`*.Authentication.Core` assembly with version and path to
help pin this down.

**Repeat sign-ins are slow / prompt for MFA every launch.** `Connect-MgGraph`
uses `-ContextScope CurrentUser` so its token cache persists to disk
across the fresh process this tool relaunches into on every start; only
an actually-expired session needs a new interactive sign-in.
`Connect-ExchangeOnline` doesn't need an equivalent flag - its cache
persists by default.

## Developer notes

**Branding.** Colors, the sidebar wordmark, and the heading typeface come
from Asurgent's CloudOps design system (extracted from the Identity
leveranspaket spec): navy sidebar gradient (`#00045A` → `#020038`), accent
blue `#2962FF`, and IBM Plex Serif/Georgia headings over an Inter/Segoe UI
body font. All brand hex values are used as-is except where the source
relies on CSS features WPF doesn't have - most notably `color-mix()` and
`rgba()` tints (the light backgrounds behind info/warning text) are
approximated as flat hex, and the badge teal (`#12A594`) is darkened to
`#0A6E62` for AA text contrast on white, since the brand's own usage puts
that teal on a tinted background rather than stark white. Neither Inter
nor IBM Plex Serif is embedded - no build step, so no font files to ship -
`FontFamily` is set to the brand font first with a system font as fallback
(`Inter, Segoe UI` / `IBM Plex Serif, Georgia`), which is silently ignored
if the brand font isn't installed rather than failing.

**Why WPF instead of a web UI.** A WPF window runs from a plain `.ps1`
with no extra runtime, browser, or local web server - important given the
distribute-to-colleagues requirement. It also looks better than default
WinForms without asking a PowerShell-only maintainer to learn a second
language.

**Every callback uses `.GetNewClosure()`.** PowerShell resolves a
variable or function referenced inside a scriptblock by walking the scope
chain active **when the scriptblock runs**, not where it was written.
Two unrelated scriptblocks using the same variable name can collide the
moment one is invoked from inside the other's function.
`.GetNewClosure()` snapshots free variables at creation time, making a
scriptblock safe to hand off and invoke from anywhere. Two things it does
**not** do:

- Protect a `function` defined inside another scriptblock - that function
  is invisible once invoked from a different scope. This tool never
  nests a `function` inside an event handler; reusable handler logic is
  always a `.GetNewClosure()`'d scriptblock variable instead (see
  `$AddLog` in `Start-MeetingRoomProvisioning.ps1`).
- Reliably capture an explicitly scope-qualified reference like
  `$Script:State.Password` when it sits inside a scriptblock built *while
  already running inside another closure* (e.g. a per-room `-Action`
  scriptblock built inside the `Add_Click` handler). The fix is to copy
  the value into a plain local (`$password = $Script:State.Password`)
  before building the inner closure, and reference the plain local inside
  it - plain free variables close over correctly no matter how many
  closures deep; only the explicit scope-qualifier trips this up.
- Persist a mutation to a captured **plain-value** variable (int, string,
  bool) across separate invocations of the *same* closure instance - each
  `& $theClosure` call re-increments from the snapshot taken when
  `.GetNewClosure()` ran, so e.g. `$counter++` inside a callback invoked
  once per module reads back `1` on every single call instead of
  accumulating (confirmed by reproducing it in isolation - see
  `Install-RoomProvisioningModules`'s progress counter in
  `RoomProvisioning.Connections.psm1`). The fix is the same shape as
  `$Script:CancelState` elsewhere in this file: capture a **reference
  type** (a hashtable) instead, and mutate a field on it
  (`$counter.Value++`) - the closure still only snapshots the reference
  once, but the object it points to is genuinely shared and its mutations
  persist normally.

## Known limitations

- The GUI runs on the UI thread with no background runspaces, so it looks
  briefly unresponsive during module install and sign-in (status text
  still updates between blocking calls). Retry waits are the exception -
  chunked and dispatcher-pumped so Cancel stays responsive. A reasonable
  trade-off for an internal admin tool against the complexity of a fully
  async WPF/runspace setup.
- License display names come from a small built-in table of common SKUs
  (`RoomProvisioning.Graph.psm1`); anything else falls back to the raw
  SKU part number.
- SSPR's actual group scope can't be read or set via Graph at all (see
  [SSPR exclusion](#sspr-exclusion)) - a Microsoft platform gap, so that
  one step always stays a manual portal action.
