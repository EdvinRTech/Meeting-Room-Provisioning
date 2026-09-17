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

If you still get `Authorization_RequestDenied` despite genuinely holding
Global Administrator, the Connect step's success message now also shows
which account actually signed in and that account's **currently active**
directory roles (queried live via Graph, not just displayed from memory).
This matters specifically for tenants using PIM (Privileged Identity
Management): a Global Administrator assignment that's *eligible* rather
than *activated* for the current sign-in does not appear in that list and
does not carry the role's permissions for this session, even though the
Entra portal still shows the person holding the role. If Global
Administrator (or whichever role you expect) is missing from that list,
activate it for this session (or sign in with an account where it's
already active) and reconnect.

**The actual root cause, found after ruling out everything above:**
`Authorization_RequestDenied` on the password step persisted even with
Global Administrator confirmed active, the right scope confirmed
requested, and a guaranteed-fresh (non-cached) token - because none of
that guarantees the interactive consent for this tool's specific *set* of
Graph permissions was ever fully, properly recorded server-side for this
app in this tenant. `(Get-MgContext).Scopes` can list a scope as
"requested" even when the real admin-consent grant behind it never fully
completed - which is apparently what happened here: signing in via WAM
for this tool's six-scope request never showed (or didn't complete) a
proper consent screen, while a separate, narrower-scoped script *did*
show one and completed it - and because Graph consent is recorded per
app + tenant, not per script, that fixed it for this tool too.

To stop this from silently recurring for the next tenant/user, Connect-
RoomProvisioningServices now calls `Test-RequiredGraphScopesGranted`
(`RoomProvisioning.Connections.psm1`) right after connecting: it queries
the tenant's actual OAuth2 permission grant for this app via
`Get-MgOauth2PermissionGrant` (the real server-side record) instead of
trusting the session's own reported scope list, and fails immediately
with a clear message naming exactly which permission is missing and
where to fix it (Entra admin center → Enterprise applications → the app
→ Permissions → "Grant admin consent") - rather than letting you discover
it as a cryptic 403 during whatever step happens to need that permission,
possibly much later. Verified against the real tenant: the query
correctly confirmed all six required scopes as genuinely granted once the
underlying consent issue was actually fixed.

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

## Module installation: fixed versions, elevated, clean every run

Clicking **Connect** first disconnects any existing Exchange Online /
Graph sessions, then removes *every* existing install of
`ExchangeOnlineManagement` and each `Microsoft.Graph.*` submodule the
tool uses - both CurrentUser and AllUsers scope. The disconnect step
matters specifically because an active session can hold those module
files open, which is the most common reason `Uninstall-Module` fails
with "module is in use". Elevation matters too: removing an AllUsers-
scope install needs admin rights, and without that, an old/wrong version
left behind on disk could still get loaded instead of the one this tool
is about to install, causing hard-to-diagnose assembly-version errors
that have nothing to do with what you're actually trying to do.

All five `Microsoft.Graph.*` submodules are then installed at one fixed
version (`$Script:GraphModuleVersion` in `RoomProvisioning.
Connections.psm1`, currently `2.40.0`), since the Microsoft.Graph SDK's
submodules only work correctly together when their versions match. This
used to be worked out live by querying PSGallery for each submodule's
latest version and picking the newest one common to all of them - that
added several network round-trips' worth of delay to every Connect for
comparatively little benefit, so it was simplified to a fixed, known-good
version. There's no automatic re-check anymore; bump the version by hand
if a future release is needed (e.g. a security fix).

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

**Not username/password.** Microsoft Graph's `Connect-MgGraph` has no
username+password parameter for interactive sign-in at all (only device
code, browser/WAM, or certificate/app-only), and Exchange Online has
largely retired password-only auth tenant-wide for security reasons - a
properly MFA-protected admin account couldn't use it anyway. There's no
version of "paste your password" that actually works here for either
service, so this tool doesn't attempt it.

`Connect-MgGraph` is called with `-ContextScope CurrentUser` specifically
so repeat sign-ins are fast: its default (`-ContextScope Process`) only
keeps the signed-in token valid for the current process, and this script
relaunches itself into a brand-new process on *every* launch (for
elevation/STA/PowerShell 7 - see above), which otherwise means a full
interactive sign-in, MFA included, every single time you start the tool,
even seconds after the last one. `CurrentUser` persists the token cache
to disk so a still-valid sign-in from a previous launch is reused
silently - no popup, no MFA - and only an actually-expired session needs
a fresh interactive sign-in. `Connect-ExchangeOnline` doesn't need an
equivalent flag; its own token cache already persists across sessions by
default.

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
   is". Country is a dropdown of every country's proper English name,
   not a free-text field - `Set-Place -CountryOrRegion` actually wants a
   2-letter ISO code (`SE` for Sweden), which a plain text box gave no
   hint about; the dropdown's value is the code, its label is the name,
   built from .NET's own region data (`System.Globalization.RegionInfo`)
   rather than a hand-typed list, so it's complete with no maintenance.
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
   the end **only for rooms it was actually confirmed set on** - an
   earlier version of this summary always claimed the password was set
   whenever the step ran, even after every retry attempt had failed; it
   now tracks success/failure per room and says so explicitly (all
   succeeded / all failed / which specific rooms failed), and the result
   card turns amber instead of green when anything failed.

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
an unrelated `$log` inside a helper function it called (since simplified
away along with the live PSGallery version lookup it supported - see
"Module installation" above), and the error surfaced as `Method
invocation failed because [System.Management.Automation.ScriptBlock]
does not contain a method named 'Add'` - which points nowhere near the
actual cause.

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

## Run log is colored by outcome

Each line in the Review & Create/Apply step's run log is colored based on
what it says, not tracked separately as state: anything containing
"FAILED" or "Cancelled" is red, recognized completions ("created",
"configured", "added", "set after", "installed", "removed", "processed",
"connected", "applied", "granted") are green, and everything else
(section headers, "Using existing X", retry-attempt progress) is cyan.
This meant switching `txtLog` from a plain `.Text` string to WPF's
`Inlines` (a `Run` per line with its own `Foreground`, plus a
`LineBreak`) - a single `TextBlock.Text` can only be one color for its
entire contents, so per-line coloring needs the richer inline-content
model instead.

## Exit button replaces Next on the last step

The last step's "Next" button turns into "Exit" in the same spot rather
than disappearing, so there's always something in that corner instead of
a step that just ends. Clicking it asks "Are you sure you want to exit?"
first, the same pattern as the Cancel button, so an out-of-habit extra
click on what used to be "Next" doesn't close the app by surprise.

## Cancelling a run in progress

The Review & Create/Apply step shows a **Cancel** button once a run
starts (next to Create/Apply). Clicking it asks for confirmation first
("The room currently being processed may be left partially configured -
some settings applied, others not. Rooms already fully processed keep
whatever was done to them.") before actually stopping anything, so an
accidental click can't cut off a run that would otherwise have finished
fine.

This exists specifically for the case where a step is going to fail
*every* retry no matter how long you wait (e.g. a genuine permissions
problem - 10 attempts at 20 seconds apart is over 3 minutes of a
guaranteed failure) and there was previously no way to stop early short
of killing the whole app. Making that possible needed more than just
adding a button: this tool has no background thread, so a plain
`Start-Sleep` during a retry wait freezes the *entire window* - no click
of any kind gets processed until the sleep ends. `Invoke-WithRetryProgress`
(`RoomProvisioning.Common.psm1`) now cuts each wait into ~200ms chunks
and pumps the WPF dispatcher between them (the same mechanism `Sync-UI`
already used to keep status text updating during long calls) - that's
what lets a Cancel click actually get processed and its handler run
*while* a retry is "sleeping", typically stopping the operation within
about one chunk instead of only between whole attempts. Verified with a
real WPF window: a click fired asynchronously partway through what would
otherwise be a 10-second wait was processed and interrupted the wait in
about 1.4 seconds.

Cancelling is checked at the start of each room and inside each retry
wait, so it takes effect promptly, but a step already talking to
Graph/Exchange (the actual network call, not the wait around it) always
finishes that one call first - there's no way to safely interrupt a
request already in flight.

## Known limitations

- The GUI runs everything on the UI thread (no background runspaces), so
  it will look briefly unresponsive during module install and sign-in -
  status text still updates between those blocking calls. Retry waits are
  the exception: they're chunked and dispatcher-pumped specifically so the
  Cancel button (see above) stays responsive during them. For an internal
  admin tool run a handful of times per tenant, this was a reasonable
  trade-off against the complexity of a fully async WPF/runspace setup.
- License display names are matched from a small built-in table of common
  SKUs (`RoomProvisioning.Graph.psm1`); anything not in that table falls
  back to showing the raw SKU part number, which is still meaningful to
  an admin.
