# Meeting Room Provisioning Tool

A GUI wizard that replaces `ExchangeOnline/New-MTRAutomatedGUI.ps1`'s
hardcoded, single-tenant script with a dynamic, per-tenant tool: pick or
create a Room List, pick or create the Conditional Access exclusion group,
type in room names, set place info, choose calendar processing rules in
plain language, and create everything in one run.

## Required admin roles for whoever signs in

Signing in with only **Exchange Administrator** is not enough - the
Connect step will work, and most of the wizard will too, but the final
password-setting step will fail with `Authorization_RequestDenied`.
Setting a user's password via Microsoft Graph (`Update-MgUser
-PasswordProfile`, used for the room's password) requires the signed-in
account to hold a privileged Entra directory role in addition to the
Graph API permission - the permission grant alone isn't enough, by
Microsoft's design. **Password Administrator** is the narrowest role
that covers it (scoped to resetting passwords for non-admin accounts,
which room mailboxes are); **User Administrator** or **Global
Administrator** also work if that's what's already in place. Sign in
with an account that holds **Exchange Administrator + one of those
password-capable roles** (or Global Administrator, which covers
everything on its own).

This was confirmed end-to-end against a real test tenant using app-only
(certificate) auth with only Exchange Administrator assigned: every step
up through adding the room to the security group succeeded, and the
password step failed with exactly this error - adding Password
Administrator to the same identity fixed it with no code changes needed.

## Running it

```powershell
.\Start-MeetingRoomProvisioning.ps1
```

Double-clicking the file (Run with PowerShell) also works. On every
launch, the script relaunches itself once more into a single consistent
state - elevated (Administrator), STA mode (required for the GUI), and
under PowerShell 7 if available (see below) - before anything else runs.
**Expect a UAC prompt every time you launch this tool**; that's Windows
asking you to consent to the elevation, not something this script can or
should hide. Say yes to it.

It installs required PowerShell modules on first Connect (see "Module
installation" below) and signs you in to both Exchange Online and
Microsoft Graph.

## Module installation: matched versions, elevated, clean every run

Clicking **Connect** first removes *every* existing install of
`ExchangeOnlineManagement` and each `Microsoft.Graph.*` submodule the
tool uses - both CurrentUser and AllUsers scope. This is why the script
relaunches itself elevated on every launch: removing an AllUsers-scope
install needs admin rights, and without that, an old/wrong version left
behind on disk could still get loaded instead of the one this tool is
about to install, causing hard-to-diagnose assembly-version errors that
have nothing to do with what you're actually trying to do.

It then works out one version that is actually published for *every*
Graph submodule (the newest version common to all of them - see
`Get-MatchedGraphModuleVersion` in `Modules\RoomProvisioning.Connections.
psm1`) and installs that exact version for each one, since the
Microsoft.Graph SDK's submodules only work correctly together when their
versions match.

`ExchangeOnlineManagement` is pinned to **3.6.0** specifically
(`$Script:ExchangeOnlineManagementVersion` in `RoomProvisioning.
Connections.psm1`), not "latest": newer versions default interactive
sign-in to the Windows account broker (WAM), which can silently sign in
with whatever Windows account is already logged in on the PC instead of
prompting for the admin account being typed in - if that Windows account
isn't licensed/enabled for Exchange Online, sign-in fails with
`AADSTS500014` ("service principal ... is disabled") even though the
intended admin account is completely fine. 3.6.0 predates that default
and reliably prompts for the account you actually want to sign in with.

Every module, once installed, is imported by its exact installed path
(queried back from `Get-InstalledModule`), not by name+version search -
belt-and-suspenders on top of the clean uninstall, since a Graph
submodule's manifest can separately trigger an internal, unpinned load of
*another* submodule by name, and importing by exact path removes any
ambiguity about which physical file that resolves to.

If you still see an assembly-loading error after Connect, it may mean
another product on that machine (e.g. an `Az.*` module, which also ships
its own `Azure.Core`) is installed and its assemblies get loaded first -
that's outside what this tool's own module management can control. The
Connect step's error display includes a list of every loaded
`Azure.Core`/`*.Authentication.Core` assembly with its version and file
path specifically to help pin down a case like that.

## CurrentUser modules are still preferred on top of all that

`$env:PSModulePath` is also **reordered** (for this process only -
nothing changes system-wide or for other PowerShell windows) so
CurrentUser-scope paths come first, as extra insurance alongside the
clean-uninstall-then-install approach above.

## This tool relaunches itself under PowerShell 7, not Windows PowerShell

The Microsoft Graph PowerShell SDK ships a separate "Desktop" build of
`Azure.Core` specifically for Windows PowerShell 5.1's .NET Framework
runtime, and that build has a confirmed incompatibility with recent SDK
releases' `Authentication.Core` - it throws *"Method GetTokenAsync ...
lacks an implementation"* the moment `Connect-MgGraph` tries to sign in.
This was confirmed with the diagnostic in the Connect error display: the
conflicting `Azure.Core.dll` and `Microsoft.Graph.Authentication.Core.dll`
both came from the exact same single install (ruling out every
version-mismatch-between-copies explanation this tool's other fixes were
built around) - `...\Microsoft.Graph.Authentication\<version>\Dependencies
\Desktop\Azure.Core.dll` is the tell. PowerShell 7's .NET (Core) build of
`Azure.Core` doesn't have this bug.

So `Start-MeetingRoomProvisioning.ps1`'s STA-relaunch logic, right at the
top of the file, now also actively prefers `pwsh.exe` over whatever
launched it: even if you start the script from Windows PowerShell
(double-click, "Run with PowerShell"), it detects that, finds PowerShell
7 if it's installed, and relaunches itself through that instead. If
PowerShell 7 isn't installed at all, it falls back to Windows PowerShell
and the GUI shows a warning on the Connect step recommending you install
it (`winget install Microsoft.PowerShell`) - Graph sign-in will likely
keep failing with the error above until you do.

## Signing in uses the normal Windows account broker (WAM) popup

Both `Connect-ExchangeOnline` and `Connect-MgGraph` sign in with the
default interactive flow - Windows' account broker (WAM), the same
native sign-in window you'd get from any Microsoft 365 app. This tool
briefly used `-UseDeviceCode` for Graph instead, to work around a bug
that only existed when running under Windows PowerShell 5.1 (a
Desktop-CLR build of `Azure.Core` incompatible with the Graph SDK - see
the PowerShell 7 section above). Now that this tool always runs under
PowerShell 7, that bug doesn't apply and the normal WAM popup works.

## No console windows stay open

Every console window this tool opens is hidden by default
(`Initialize-ConsoleVisibilityControl` / `Hide-RoomProvisioningConsole`
in `RoomProvisioning.Connections.psm1`) - both sign-in flows use their
own native popup window, not anything printed to the console, so there's
nothing to show in it at any point. The relaunch this script does at
startup (to get an STA thread / elevate / prefer PowerShell 7 - see
above) also starts hidden and doesn't linger: it doesn't `-Wait`, so its
own brief window closes the instant the real GUI process is started.
That GUI process's console and its WPF window are the same process, so
closing the GUI window closes everything - no separate window to clean
up afterward.

This is all gated behind a `-RelaunchedForGui` switch that's only ever
set automatically by the script's own relaunch - running the script
directly from your own terminal (for development/testing) leaves your
terminal alone entirely.

## Distributing to colleagues

Copy the whole `MeetingRoomProvisioning` folder (it needs `Start-
MeetingRoomProvisioning.ps1`, `UI\MainWindow.xaml`, and everything under
`Modules\`) - there's nothing to install beyond what the tool installs
for itself on first run. No separate runtime, no build step.

If a colleague's execution policy blocks the script, they can run:

```powershell
Unblock-File .\Start-MeetingRoomProvisioning.ps1
```

## What each step does

1. **Connect** - signs in once to Exchange Online and once to Microsoft
   Graph (scopes: `User.ReadWrite.All`, `Group.ReadWrite.All`,
   `Policy.Read.All`, `Policy.ReadWrite.ConditionalAccess`,
   `Directory.Read.All`, `Organization.Read.All`). Also checks every
   *existing* room mailbox for an assigned license and shows a summary -
   informational only, this tool never assigns a license itself.
2. **Mode** - **Create new rooms** (the original flow) or **Edit existing
   rooms** (change calendar processing, place info, Room List membership,
   and/or password on rooms that already exist). This choice changes
   which of the steps below apply and how they behave - see "Edit
   existing rooms" below for specifics.
3. **Room List** - queried dynamically (`Get-DistributionGroup
   -RecipientTypeDetails RoomList`); pick one, create a new one (just
   type a name - its email address is generated automatically from the
   tenant's default domain, no separate address field), or in Edit mode,
   leave membership unchanged.
4. **Conditional Access exclusion group** *(Create mode only - skipped
   entirely in Edit mode)* - queried dynamically by scanning every CA
   policy's excluded groups; pick one or create a new one. A new group is
   automatically excluded from every existing CA policy via Graph.
5. **Rooms** - Create mode: type room names (Enter/Add, repeat) and pick
   a domain (`Get-MgDomain`, verified domains only). Edit mode: pick one
   or more existing room mailboxes from a list instead.
6. **Place info** - maps to `Set-Place`. Any field left blank is left out
   of the command entirely (not passed as empty) - for an existing room
   in Edit mode, a blank field simply means "leave this as it already
   is".
7. **Calendar processing** - Standard mode uses the same defaults the
   original script always applied. Custom mode asks plain-language
   questions and translates them to `Set-CalendarProcessing` parameters
   (see "Calendar processing cleanup" below). In Edit mode there's also a
   "don't change calendar processing" option, selected by default.
8. **Review & create/apply** - shows a summary, then in Create mode
   creates each room mailbox (skips ones that already exist), adds it to
   the Room List, applies calendar processing and place info, then adds
   it to the security group and sets its password. In Edit mode, it
   applies whichever of Room List / calendar processing / place info you
   changed to each selected room, and resets the password only if you
   checked "Reset password for these rooms" (unchecked by default - an
   edit run doesn't touch the password unless you ask it to). Group
   membership and password-setting retry (default: 10 attempts, 20s
   apart) because a just-created or just-changed account isn't always
   immediately visible to Graph/Exchange writes - place info retries too,
   for the same reason (confirmed live: a room's very first `Set-Place`
   call can fail with `PlaceNotFoundInDirectory` moments after the
   mailbox is created). The progress bar's max is the retry cap, so it
   reflects how many attempts are actually left rather than spinning
   generically. The shared password (`REDACTED-ROTATE-THIS-PASSWORD`) is displayed at
   the end when it was set.

## Edit existing rooms

A second mode alongside room creation, for changing settings on rooms
that already exist rather than provisioning new ones. Deliberately
narrower in scope than Create mode:

- **No mailbox creation** and **no Conditional Access / security group
  step** - both are Create-mode-only; picking Edit mode skips that step
  in the wizard entirely (Back/Next jump over it).
- **Room List, calendar processing, and password are all optional** -
  each defaults to "don't change" (Room List and calendar processing via
  a dedicated radio option; password via an unchecked "Reset password
  for these rooms" checkbox on the review step). Nothing you don't
  explicitly opt into gets touched.
- **Place info always applies**, but blank fields are omitted from the
  `Set-Place` call exactly like in Create mode - for an existing room
  that means "leave this field as it already is", not "clear it".
- Rooms are picked from a live list (`Get-ExistingRoomMailboxes` in
  `RoomProvisioning.Exchange.psm1`) with multi-select, so one run can
  apply the same change to several rooms at once.

Verified end-to-end against a real tenant: created two rooms, confirmed
they appear back in the existing-rooms list, then applied a second,
different calendar-processing configuration, different place info, and a
password reset to both - and confirmed with `Get-Place`/
`Get-CalendarProcessing` afterward that the changes actually took effect,
not just that the commands didn't error.

## Calendar processing cleanup

The original script's private-room example set `AddOrganizerToSubject
$false` alongside `DeleteSubject $true` / `DeleteComments $true`, but
`AddOrganizerToSubject $false` was already the script's *default* for
every room - so it wasn't actually a privacy-specific setting. In this
tool, `AddOrganizerToSubject` always stays `$false` (a Teams Rooms panel
has no use for the organizer's name in the subject line either way), and
the "private room" question only toggles `DeleteSubject`/`DeleteComments`.

Custom mode also exposes a few more commonly-requested settings beyond
what the spec listed: recurring meetings allowed, attachment stripping,
external meeting requests, removing the "Private" flag on synced
meetings, a custom auto-reply message, and delegate-approval bookings
(`AllBookInPolicy`/`AllRequestOutOfPolicy` off + `ResourceDelegates`).

## Why WPF instead of a web UI

Kept everything in PowerShell + XAML (XAML only describes the layout, no
logic) rather than an HTML/JS front end, mainly because of the
distribute-to-colleagues requirement: a WPF window runs from a plain
`.ps1` with no extra runtime, browser, or local web server to stand up.
Looks noticeably better than default WinForms (styled buttons, modern
color palette, resizable layout) without asking a PowerShell-only admin
to learn a second language to maintain it later.

## Why every callback uses GetNewClosure

Every scriptblock in this tool that gets stored in a variable, passed
into a module function as a `-ProgressCallback`/`-LogCallback`/`-Action`,
or attached to a WPF event, ends with `.GetNewClosure()`. This isn't
stylistic - without it, the tool intermittently breaks in a way that's
very hard to diagnose.

PowerShell resolves a variable or function name referenced inside a
`{ ... }` scriptblock by walking the scope chain **active at the moment
the scriptblock is invoked**, not the scope where it was written. Two
scriptblocks that both use a variable named e.g. `$log`, one holding a
`List[string]` and another holding something else entirely, can collide
the moment one is invoked from inside the other's function - the inner
one "sees" whichever `$log` happens to be in scope at the call site, not
the one its author meant. This is exactly what happened during
development: a `$log` in `Install-RoomProvisioningModules` collided with
an unrelated `$log` inside `Get-MatchedGraphModuleVersion`, and the
error surfaced as `Method invocation failed because
[System.Management.Automation.ScriptBlock] does not contain a method
named 'Add'` - which points nowhere near the actual cause.

`.GetNewClosure()` snapshots the scriptblock's free variables at the
point it's created, making it behave the way you'd naturally expect -
safe to hand off to another function and invoke from anywhere. It does
**not** protect a `function` defined inside another scriptblock (that
function is invisible once invoked from a different scope entirely), so
this tool never defines a nested `function` inside an event handler -
every reusable piece of handler logic is a `.GetNewClosure()`'d
scriptblock variable instead (see `$AddLog` in
`Start-MeetingRoomProvisioning.ps1` for the pattern).

It also does **not** reliably protect an explicitly scope-qualified
reference like `$Script:State.Password` when that reference sits inside
a scriptblock created (and `.GetNewClosure()`'d) *while already running
inside another closure* - e.g. building a per-room `-Action` scriptblock
inside the `Add_Click` handler, which is itself a closure. That exact
pattern surfaced as `Cannot bind argument to parameter 'Password'
because it is an empty string` even though the review screen showed the
password correctly - confirmed by reproducing it in isolation outside
this project entirely. The fix: capture the value into a plain local
variable (`$password = $Script:State.Password`) *before* building the
inner closure, and reference that plain local inside it instead of the
`$Script:`-qualified path - plain free variables close over correctly no
matter how many closures deep, only the explicit scope-qualifier trips
this up. See `$password` right before the room-creation loop in
`Start-MeetingRoomProvisioning.ps1`.

## Architecture

```
Start-MeetingRoomProvisioning.ps1   Entry point: loads XAML, wires events, orchestrates creation
UI/MainWindow.xaml                  Window layout only, no logic
Modules/
  RoomProvisioning.Common.psm1      Generic retry-with-progress helper
  RoomProvisioning.Connections.psm1 Module install + EXO/Graph sign-in
  RoomProvisioning.Exchange.psm1    Room List, mailbox creation, Set-Place
  RoomProvisioning.Graph.psm1       CA policy exclusion group, domains, license check, password
  RoomProvisioning.CalendarLogic.psm1  Plain-language -> Set-CalendarProcessing mapping
```

## Known limitations

- The GUI runs everything on the UI thread (no background runspaces), so
  it will look briefly unresponsive during module install, sign-in, and
  each retry wait - status text still updates between those blocking
  calls. For an internal admin tool run a handful of times per tenant,
  this was a reasonable trade-off against the complexity of a fully async
  WPF/runspace setup.
- License display names are matched from a small built-in table of common
  SKUs (`RoomProvisioning.Graph.psm1`); anything not in that table falls
  back to showing the raw SKU part number, which is still meaningful to
  an admin.
