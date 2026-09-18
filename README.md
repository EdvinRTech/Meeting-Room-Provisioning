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
  PowerShell 7 if installed - see [Troubleshooting](#troubleshooting)).
  Install PS7 with `winget install Microsoft.PowerShell` if you don't have
  it; the tool will still run without it, but Graph sign-in is likely to
  fail (see below).
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

```powershell
.\Start-MeetingRoomProvisioning.ps1
```

Double-clicking the file (Run with PowerShell) also works. Expect a UAC
prompt on every launch - the tool relaunches itself elevated, in STA mode,
and under PowerShell 7 if available, before anything else runs.

The first Connect click installs required modules (Exchange Online
Management, Microsoft.Graph submodules) and signs you in to both Exchange
Online and Microsoft Graph via the normal Windows sign-in popup (WAM) -
there's no username/password field, since neither service supports plain
password auth for this kind of sign-in anymore. A progress bar advances
through the whole Connect step's six phases (disconnect existing
sessions, remove old module versions, install, import, sign in, load
tenant data), alongside short, plain-language status text ("Removing old
module versions...", "Installing modules..."), instead of just spinning
generically. Module install/removal failures are still recorded in full
detail and surfaced if the step actually fails, but per-module success/failure lines
(which can include a raw, sometimes multi-sentence .NET exception message)
are no longer flashed past one at a time on that single status line.

### Distributing to colleagues

Copy the whole `MeetingRoomProvisioning` folder (`Start-MeetingRoomProvisioning.ps1`,
`UI\MainWindow.xaml`, and everything under `Modules\`) - nothing else to
install, no build step. If execution policy blocks it:

```powershell
Unblock-File .\Start-MeetingRoomProvisioning.ps1
```

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
the wizard and the final result both say so.

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

Password-setting, group membership, place-info, and SSPR-exclusion calls
all retry (10 attempts, 20s apart by default) because a just-created or
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
Start-MeetingRoomProvisioning.ps1   Entry point: loads XAML, wires events, orchestrates creation
UI/MainWindow.xaml                  Window layout only, no logic
Modules/
  RoomProvisioning.Common.psm1      Generic retry-with-progress helper
  RoomProvisioning.Connections.psm1 Module install + EXO/Graph sign-in
  RoomProvisioning.Exchange.psm1    Room List, mailbox creation, Set-Place
  RoomProvisioning.Graph.psm1       CA policy exclusion group, SSPR exclusion, domains, license check, password
  RoomProvisioning.CalendarLogic.psm1  Plain-language -> Set-CalendarProcessing mapping
```

## Troubleshooting

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
whatever launched it for exactly this reason; if PowerShell 7 isn't
installed, it falls back to Windows PowerShell and the Connect step shows
a warning recommending `winget install Microsoft.PowerShell`.

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
