# Meeting Room Provisioning Tool

A GUI wizard that replaces `ExchangeOnline/New-MTRAutomatedGUI.ps1`'s
hardcoded, single-tenant script with a dynamic, per-tenant tool: pick or
create a Room List, pick or create the Conditional Access exclusion group,
type in room names, set place info, choose calendar processing rules in
plain language, and create everything in one run.

## Running it

```powershell
.\Start-MeetingRoomProvisioning.ps1
```

Double-clicking the file (Run with PowerShell) also works. The script
relaunches itself in STA mode automatically if needed (required for the
GUI to display), installs required PowerShell modules on first Connect
(see "Module installation is intentionally invasive" below), and signs
you in to both Exchange Online and Microsoft Graph.

## Module installation: matched versions, CurrentUser scope, no uninstalling

Clicking **Connect** installs `ExchangeOnlineManagement` and every
`Microsoft.Graph.*` submodule the tool uses to `-Scope CurrentUser` (no
admin rights needed). For the Graph submodules specifically, it first
works out one version that is actually published for *all* of them (the
newest version common to every submodule - see
`Get-MatchedGraphModuleVersion` in `Modules\RoomProvisioning.Connections.
psm1`) and pins every submodule to that exact version, since the
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

An earlier version of this tool also force-uninstalled every existing
install first. That was dropped: on a machine where the existing copies
are AllUsers-scoped (`Program Files\...`), removing them needs admin
rights this tool deliberately doesn't require, so the uninstall step
could never actually succeed there - it just produced a failed-removal
message on every run without changing anything. Reliability instead
comes from installing the matched version to CurrentUser scope and
always importing it by its exact installed path (see the next section)
rather than depending on any existing copy being gone.

If you still see an assembly-loading error after Connect, it may mean
another product on that machine (e.g. an `Az.*` module, which also ships
its own `Azure.Core`) is installed and its assemblies get loaded first -
that's outside what this tool's own module management can control. The
Connect step's error display includes a list of every loaded
`Azure.Core`/`*.Authentication.Core` assembly with its version and file
path specifically to help pin down a case like that.

## CurrentUser modules are preferred, and imported unambiguously

Right at the top of `Start-MeetingRoomProvisioning.ps1`, before anything
else runs, `$env:PSModulePath` is **reordered** (for this process only -
nothing changes system-wide or for other PowerShell windows) so
CurrentUser-scope paths come first. Nothing is removed: an earlier
version of this fix removed the AllUsers paths outright, which broke
`Get-InstalledModule` on machines where PowerShellGet itself is only
installed AllUsers rather than under the built-in system path - a
straightforward reorder avoids that regression while still fixing the
original problem.

On top of that, `Install-RoomProvisioningModules` never imports the
modules it just installed by name+version search - it asks
`Get-InstalledModule` for the exact `InstalledLocation` it just installed
to and imports that `.psd1` file directly. Name+version search still
walks `$env:PSModulePath`, and a Graph submodule's manifest can trigger
an internal, unpinned load of *another* submodule by name as a side
effect - if a stale AllUsers copy (which this tool can't remove without
admin rights) got found along the way, .NET would end up with two
physically different DLL builds of the same type loaded at once. That
mismatch is what caused errors like *"Method GetTokenAsync ... lacks an
implementation"* - nothing to do with how sign-in authenticates, purely
the wrong assembly getting loaded. Importing by exact file path removes
that ambiguity for this tool's own top-level imports entirely, and the
path reorder covers the internal-dependency-load case too.

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

## Signing in uses device code, not the default popup

`Connect-MgGraph` normally tries to sign in using Windows' account broker
(WAM) - an embedded, native sign-in window. That broker component is
inconsistently present/working across machines, especially VMs. So
`Connect-RoomProvisioningServices` always signs in to Graph with
`-UseDeviceCode` instead: it prints a one-time code and
`https://microsoft.com/devicelogin`, and you finish signing in in your
normal web browser. Exchange Online still uses its own regular sign-in
popup (a separate window, unrelated to the console).

## No console windows stay open

Every console window this tool opens is hidden by default - there's
nothing to look at in it except during the one moment
`Connect-RoomProvisioningServices` needs to show you the Graph
device-sign-in code, when it un-hides its own window just long enough
for that, then hides it again (`Show-RoomProvisioningConsole` /
`Hide-RoomProvisioningConsole` in `RoomProvisioning.Connections.psm1`).
The relaunch this script does at startup (to get an STA thread / prefer
PowerShell 7 - see above) also starts hidden and doesn't linger: it
doesn't `-Wait`, so its own brief window closes the instant the real GUI
process is started. That GUI process's console and its WPF window are
the same process, so closing the GUI window closes everything - no
separate window to clean up afterward.

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
2. **Room List** - queried dynamically (`Get-DistributionGroup
   -RecipientTypeDetails RoomList`); pick one or create a new one.
3. **Conditional Access exclusion group** - queried dynamically by
   scanning every CA policy's excluded groups; pick one or create a new
   one. A new group is automatically excluded from every existing CA
   policy via Graph.
4. **Room names & domain** - type a name, press Enter/Add, repeat. Domain
   list is pulled from `Get-MgDomain` (verified domains only).
5. **Place info** - maps to `Set-Place`. Any field left blank is left out
   of the command entirely (not passed as empty).
6. **Calendar processing** - Standard mode uses the same defaults the
   original script always applied. Custom mode asks plain-language
   questions and translates them to `Set-CalendarProcessing` parameters
   (see "Calendar processing cleanup" below).
7. **Review & create** - shows a summary, then creates each room mailbox
   (skips ones that already exist), adds it to the Room List, applies
   calendar processing and place info, then adds it to the security group
   and sets its password - both of the latter retry (default: 10 attempts,
   20s apart) because a just-created account isn't always immediately
   visible to Graph writes. The progress bar's max is the retry cap, so
   it reflects how many attempts are actually left rather than spinning
   generically. The shared password (`REDACTED-ROTATE-THIS-PASSWORD`) is displayed at
   the end for easy reference.

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
