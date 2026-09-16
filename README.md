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

## Module installation is intentionally invasive

Clicking **Connect** does not just install modules that are missing - it
**force-removes every existing installed version** of `ExchangeOnline
Management` and every `Microsoft.Graph.*` submodule the tool uses, then
reinstalls them from scratch. For the Graph submodules specifically, it
first works out one version that is actually published for *all* of them
(the newest version common to every submodule - see
`Get-MatchedGraphModuleVersion` in `Modules\RoomProvisioning.Connections.
psm1`) and pins every submodule to that exact version.

Why: the Microsoft.Graph SDK ships as several submodules that only work
correctly together when their versions match, and each depends on
further assemblies (like `Azure.Core`) by exact version. Having two
versions of a submodule installed, or submodules at different versions,
is the single most common cause of errors like *"Could not load file or
assembly 'Azure.Core, Version=x.x.x.x'..."* - .NET cannot load two
different versions of the same assembly into one process, and a stale or
partially-installed module version can trigger this on some machines but
not others. Force-removing and reinstalling a matched set removes that
variable entirely, so the tool behaves the same on a brand-new VM as on
a machine with a history of other scripts installing other module
versions.

Trade-off worth knowing about: this **will remove** other versions of
these modules that other scripts on the same machine might depend on. If
that's a problem on a shared machine, consider running this tool from a
dedicated VM or user profile rather than one used for other Graph/EXO
automation.

This only covers the modules this tool itself requires - if you still
see an assembly-loading error after this runs, it likely means another
product (e.g. an `Az.*` module, which also ships its own `Azure.Core`)
is installed and gets loaded first; that's outside what this tool
manages.

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

## Signing in uses device code, not the default popup

`Connect-MgGraph` normally tries to sign in using Windows' account broker
(WAM) - an embedded, native sign-in window. That broker component is
inconsistently present/working across machines, especially VMs. So
`Connect-RoomProvisioningServices` always signs in to Graph with
`-UseDeviceCode` instead: it prints a one-time code and
`https://microsoft.com/devicelogin` to the **console window that opens
alongside this app** (not the GUI window itself - check your taskbar for
it), and you finish signing in in your normal web browser there. Exchange
Online still uses its own regular sign-in popup.

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
