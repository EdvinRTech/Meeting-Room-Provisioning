<#
.SYNOPSIS
    GUI tool for provisioning Microsoft Teams meeting rooms end to end:
    room mailbox, Room List, Conditional Access exclusion group, calendar
    processing, Set-Place info, and password - all from one wizard.

.DESCRIPTION
    What this file actually is, in plain PowerShell terms:
      - The window's *layout* (buttons, text boxes, colors) lives in
        UI\MainWindow.xaml. XAML is just a markup language (like a very
        strict, tag-based way to describe "put a button here") - there is
        no logic in it, it only describes what the window looks like.
      - This .ps1 file is 100% PowerShell. It loads that XAML file into a
        real WPF window object, finds each named control (the x:Name
        attributes in the XAML, e.g. x:Name="btnConnect"), and attaches
        normal PowerShell scriptblocks to their events (Add_Click, etc.)
        - exactly like assigning a scriptblock to a variable, except WPF
        calls it for you when the button is clicked.
      - All the actual Exchange Online / Microsoft Graph work happens in
        the .psm1 modules under .\Modules - this script just wires the
        GUI to those functions.

.NOTES
    Author:  Edvin Rodin, Claude Code
    Company: Asurgent AB
#>

param(
    # Internal - set automatically when this script relaunches itself.
    # Marks "this console window belongs to us, it's safe to hide/show
    # it programmatically" - never set this when running the script
    # directly from your own terminal, or your terminal's window could
    # get hidden along with it.
    [switch]$RelaunchedForGui
)

#========================================================#
# On every launch, this script relaunches itself once (as a hidden
# window - see below) to end up in exactly one consistent state, for a
# one-click "just run it" experience with no setup steps for whoever's
# using it:
#
#   1. Elevated (Administrator). Needed so Install-RoomProvisioningModules
#      can actually remove an existing AllUsers-scope module install
#      before reinstalling the matched/pinned versions - without
#      elevation, Uninstall-Module fails for those, an old version stays
#      on disk, and it can still get loaded instead of the one this tool
#      just installed. Windows will show its own UAC consent prompt for
#      this - that's an OS-level security prompt this script cannot
#      hide, and shouldn't try to.
#   2. STA (single-threaded apartment). WPF requires it; PowerShell 7
#      (pwsh.exe) defaults to MTA, Windows PowerShell 5.1 defaults to STA
#      already.
#   3. Running under PowerShell 7 rather than Windows PowerShell 5.1,
#      even if you started the script from Windows PowerShell
#      (double-click, "Run with PowerShell", etc.) - preferred whenever
#      pwsh.exe is installed. Reason: the Microsoft Graph PowerShell SDK
#      ships a separate "Desktop" build of Azure.Core specifically for
#      Windows PowerShell 5.1's .NET Framework runtime, and that build
#      has a confirmed incompatibility with recent SDK releases'
#      Authentication.Core - it throws "Method GetTokenAsync ... lacks an
#      implementation" the moment Connect-MgGraph tries to sign in,
#      regardless of which exact module version is installed or how
#      cleanly it was installed. PowerShell 7's .NET (Core) build of
#      Azure.Core doesn't have this bug. If pwsh.exe isn't installed at
#      all, this falls back to Windows PowerShell and the GUI shows a
#      warning recommending you install PowerShell 7.
#
# The relaunched process's console window is started hidden (-WindowStyle
# Hidden) rather than shown - there's nothing useful to look at in it
# except briefly during Graph device-code sign-in, when
# Connect-RoomProvisioningServices below un-hides it just long enough to
# show the code, then hides it again. The relauncher itself doesn't
# -Wait, so its own (momentary, hidden) window closes immediately once
# the real GUI process is started - and since that GUI process's console
# and its WPF window are the same process, closing the GUI closes both.
#========================================================#
function Test-IsAdministrator {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [System.Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-Pwsh7Path {
    <#
        PATH first (covers a Microsoft Store install, a portable zip
        someone put on PATH, etc.), then the default per-machine MSI/winget
        install location directly - needed because an installer updates the
        registry's Environment key, not any already-running process's
        in-memory $env:PATH, so a sibling process that just installed
        PowerShell 7 moments ago wouldn't be found by Get-Command alone.
    #>
    $cmd = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $defaultPath = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    if (Test-Path -LiteralPath $defaultPath) { return $defaultPath }
    return $null
}

$IsElevated = Test-IsAdministrator
$IsWindowsPowerShellDesktop = $PSVersionTable.PSEdition -ne 'Core'
$Pwsh = Get-Pwsh7Path

# Offer to install PowerShell 7 automatically so a machine that's never run
# this tool before doesn't need a separate manual "go install PowerShell 7"
# step - see the PowerShell 7 relaunch note below for why this tool needs
# it. Only offered once we're actually elevated (installing software needs
# admin rights too, so asking before that would just fail) and only when
# still running Windows PowerShell - a pwsh-hosted relaunch of this same
# script would otherwise ask again on every launch for no reason.
if ($IsElevated -and -not $Pwsh -and $IsWindowsPowerShellDesktop) {
    Add-Type -AssemblyName System.Windows.Forms
    $installChoice = [System.Windows.Forms.MessageBox]::Show(
        "PowerShell 7 isn't installed on this machine. This tool needs it for reliable Microsoft Graph sign-in - without it, sign-in can fail partway through with a cryptic error.`n`nInstall it now via winget (Microsoft's official package manager)? This only needs to happen once.",
        'Meeting Room Provisioning',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($installChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
        $winget = Get-Command -Name 'winget.exe' -ErrorAction SilentlyContinue
        if ($winget) {
            try {
                # Not hidden, unlike the relaunches below: this can take a
                # genuine 10-60+ seconds depending on the machine and
                # network, and showing winget's own real progress is more
                # honest than a window that looks hung with nothing visible
                # at all.
                $installProc = Start-Process -FilePath $winget.Source -ArgumentList @(
                    'install', '--id', 'Microsoft.PowerShell', '--source', 'winget',
                    '-e', '--silent', '--accept-source-agreements', '--accept-package-agreements'
                ) -Wait -WindowStyle Normal -PassThru -ErrorAction Stop
                $Pwsh = Get-Pwsh7Path
                if (-not $Pwsh) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "The PowerShell 7 installer finished (exit code $($installProc.ExitCode)), but pwsh.exe still couldn't be found afterward. Continuing under Windows PowerShell 5.1 for now - you can install PowerShell 7 yourself later with 'winget install Microsoft.PowerShell'.",
                        'Meeting Room Provisioning',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    ) | Out-Null
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Installing PowerShell 7 via winget failed: $($_.Exception.Message)`n`nContinuing under Windows PowerShell 5.1 for now - you can install PowerShell 7 yourself later with 'winget install Microsoft.PowerShell'.",
                    'Meeting Room Provisioning',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "winget (Windows Package Manager) isn't available on this machine, so PowerShell 7 can't be installed automatically. Continuing under Windows PowerShell 5.1 for now - install PowerShell 7 yourself from https://aka.ms/powershell-release?tag=stable when you can.",
                'Meeting Room Provisioning',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
    }
}

$NeedsRelaunch = (-not $IsElevated) -or ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') -or ($IsWindowsPowerShellDesktop -and $Pwsh)

if ($NeedsRelaunch) {
    $exe = if ($IsWindowsPowerShellDesktop -and $Pwsh) { $Pwsh } else { (Get-Process -Id $PID).Path }
    $relaunchArgs = @{
        FilePath     = $exe
        WindowStyle  = 'Hidden'
        ArgumentList = @('-NoProfile', '-STA', '-File', "`"$PSCommandPath`"", '-RelaunchedForGui')
    }
    if (-not $IsElevated) { $relaunchArgs.Verb = 'RunAs' }

    try {
        Start-Process @relaunchArgs -ErrorAction Stop
    } catch {
        # Most likely cause: the UAC prompt was cancelled. Nothing built
        # yet to show this in the GUI (we haven't even loaded the XAML),
        # so fall back to a plain message box.
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "This tool needs to run as Administrator to reliably install its required PowerShell modules. Relaunch failed or was cancelled:`n`n$($_.Exception.Message)",
            'Meeting Room Provisioning',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
    exit
}

$ErrorActionPreference = 'Stop'
$ScriptRoot = $PSScriptRoot

#========================================================#
# Put CurrentUser-scope module paths first on $env:PSModulePath, for this
# process only - nothing is removed, so PowerShellGet/Install-Module
# (which on some machines only live under an AllUsers path, not the
# built-in system one) stay fully reachable. This previously removed the
# AllUsers paths outright, which broke Get-InstalledModule on machines
# where PowerShellGet itself is only installed AllUsers - reordering
# instead of removing fixes that regression.
#
# Why reorder at all: this tool cannot remove an AllUsers-scope module
# without admin rights (Uninstall-Module fails, harmlessly logged), so a
# stale/mismatched copy can be left on disk. A Graph submodule's manifest
# can trigger an internal, unpinned load of another Graph submodule by
# name - if an AllUsers copy is found first, PowerShell can load its
# assemblies alongside our pinned CurrentUser copy, and .NET treats two
# physically different DLL builds of the same type as incompatible even
# when our own Import-Module -RequiredVersion asked for one exact
# version. Still worth doing for general reliability, but note it turned
# out NOT to be the cause of the "Method GetTokenAsync ... lacks an
# implementation" error seen during development - that one was a
# Windows-PowerShell-5.1-vs-Graph-SDK issue, see the pwsh-preferring
# relaunch logic above.
#========================================================#
$PSModulePathEntries = $env:PSModulePath -split [System.IO.Path]::PathSeparator
$CurrentUserModulePaths = $PSModulePathEntries | Where-Object { $_ -and $_.StartsWith($HOME, [System.StringComparison]::OrdinalIgnoreCase) }
$OtherModulePaths = $PSModulePathEntries | Where-Object { $_ -and -not $_.StartsWith($HOME, [System.StringComparison]::OrdinalIgnoreCase) }
$env:PSModulePath = (@($CurrentUserModulePaths) + @($OtherModulePaths)) -join [System.IO.Path]::PathSeparator

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.Common.psm1') -Force
Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.Connections.psm1') -Force
Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.Exchange.psm1') -Force
Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.Graph.psm1') -Force
Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.CalendarLogic.psm1') -Force

Initialize-ConsoleVisibilityControl -Enabled:$RelaunchedForGui

#========================================================#
# Load the window from XAML
#========================================================#
[xml]$XamlDoc = Get-Content -Path (Join-Path $ScriptRoot 'UI\MainWindow.xaml') -Raw
$XamlReader = New-Object System.Xml.XmlNodeReader $XamlDoc
$Window = [System.Windows.Markup.XamlReader]::Load($XamlReader)

$ElementNames = @(
    'lblStep1', 'lblStepMode', 'lblStep2', 'lblStep3', 'lblStepSspr', 'lblStep4', 'lblStep5', 'lblStep6', 'lblStep7',
    'Step1Panel', 'StepModePanel', 'Step2Panel', 'Step3Panel', 'StepSsprPanel', 'Step4Panel', 'Step5Panel', 'Step6Panel', 'Step7Panel',
    'btnConnect', 'progConnect', 'txtConnectStatus', 'LicenseInfoCard', 'txtLicenseInfo',
    'radModeCreate', 'radModeEdit',
    'radSkipRoomList', 'radUseExistingRoomList', 'lstRoomLists', 'radCreateNewRoomList', 'txtNewRoomListName', 'txtNewRoomListAddressPreview',
    'radUseExistingCAGroup', 'lstCAGroups', 'radCreateNewCAGroup', 'txtNewCAGroupName',
    'txtSsprStatus', 'SsprDisabledPanel', 'SsprGroupFoundPanel', 'SsprGroupMissingPanel', 'chkCreateSsprGroup',
    'CreateRoomNamingPanel', 'txtRoomNameInput', 'btnAddRoomName', 'lstRoomNames', 'btnRemoveRoomName', 'cmbDomain',
    'EditRoomSelectionPanel', 'lstExistingRooms',
    'txtBuilding', 'txtCapacity', 'txtCity', 'txtPostalCode', 'txtState', 'txtStreet', 'cmbCountry',
    'radSkipCalendar', 'radStandardMode', 'radCustomMode', 'CustomCalendarPanel',
    'chkIsPrivate', 'chkAllowConflictingSeries', 'txtConflictPercentage', 'txtMaxConflictInstances',
    'cmbBookingWindow', 'txtBookingWindowCustomDays', 'chkRequireApproval', 'txtApprovalDelegates',
    'chkAllowRecurring', 'chkRemoveAttachments', 'chkAllowExternalRequests', 'chkRemovePrivateFlag', 'txtAdditionalResponseText',
    'txtReviewHeading', 'txtReviewSub', 'txtReviewSummary', 'chkResetPassword', 'PasswordEntryPanel', 'pwdRoomPassword', 'pwdRoomPasswordConfirm', 'btnCreate', 'btnCancel', 'ProgressPanel', 'txtProgressStatus', 'progRetry',
    'txtLog', 'LogScrollViewer', 'ResultCard', 'txtResultHeading', 'txtResultPassword', 'txtResultLicenseReminder', 'txtResultSsprReminder',
    'btnBack', 'btnNext', 'txtGlobalError'
)
$ui = @{}
foreach ($name in $ElementNames) { $ui[$name] = $Window.FindName($name) }

# Order here is the actual step order shown to the user - Mode sits
# between Connect and Room List. Both the Security Group step (index 4)
# and the SSPR Exclusion step (index 5) apply to BOTH modes - a room being
# edited may never have been added to the CA exclusion group either (e.g.
# it predates this tool managing that, or the group didn't exist yet), so
# Edit mode gets the same pick-existing-or-create-new choice as Create
# mode instead of only ever touching that at room-creation time.
$StepPanels = @($ui.Step1Panel, $ui.StepModePanel, $ui.Step2Panel, $ui.Step3Panel, $ui.StepSsprPanel, $ui.Step4Panel, $ui.Step5Panel, $ui.Step6Panel, $ui.Step7Panel)
$StepLabels = @($ui.lblStep1, $ui.lblStepMode, $ui.lblStep2, $ui.lblStep3, $ui.lblStepSspr, $ui.lblStep4, $ui.lblStep5, $ui.lblStep6, $ui.lblStep7)
$TotalSteps = $StepPanels.Count
# Kept as an (empty) list rather than removed outright - Get-AdjacentVisibleStep
# below is written generically against it in case a future step ever needs
# to be Create-mode-only again.
$CreateOnlySteps = @()

$BrushConverter = New-Object System.Windows.Media.BrushConverter
function Get-Brush([string]$Hex) { $BrushConverter.ConvertFromString($Hex) }

function Get-DiagnosticErrorText {
    <#
        Builds a detailed multi-line description of a failure: the full
        exception chain (an assembly-loading error's real cause is often
        in .InnerException, not the top-level message), plus - for
        anything that smells like an assembly-version clash - which
        physical DLL file actually ended up loaded for Azure.Core and
        Microsoft.Graph.Authentication.Core, and its version/path. This
        turns a vague "it failed" screenshot into something that
        pinpoints the conflicting file directly instead of guessing.
    #>
    param([Parameter(Mandatory)]$ErrorRecord)

    $lines = [System.Collections.Generic.List[string]]::new()
    $ex = $ErrorRecord.Exception
    $depth = 0
    while ($ex) {
        $prefix = if ($depth -eq 0) { '' } else { ('  ' * $depth) + '-> ' }
        $lines.Add("$prefix$($ex.GetType().FullName): $($ex.Message)")
        $ex = $ex.InnerException
        $depth++
    }

    $suspectAssemblies = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -match 'Azure\.Core|Authentication\.Core' } |
        Sort-Object { $_.GetName().Name }
    if ($suspectAssemblies) {
        $lines.Add('Loaded assemblies that commonly cause this:')
        foreach ($asm in $suspectAssemblies) {
            $name = $asm.GetName()
            $location = try { $asm.Location } catch { '(dynamic/in-memory)' }
            $lines.Add("  $($name.Name) $($name.Version) - $location")
        }
    }

    return ($lines -join "`n")
}

# Pumps the WPF dispatcher so a status-text update becomes visible on
# screen *before* a long blocking call (Connect-*, Set-CalendarProcessing,
# retry loops, ...) continues running on this same thread. This tool does
# not use background threads/runspaces, so without this the window would
# just look frozen with stale text during any long step.
function Sync-UI {
    $Window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
}

$Script:State = [ordered]@{
    CurrentStep     = 1
    Connected       = $false
    Mode            = 'Create'   # 'Create' or 'Edit'
    RoomLists       = @()
    CAGroups        = @()
    SsprEnabled     = $false
    SsprGroupExists = $false
    Domains         = @()
    DefaultDomain   = $null
    ExistingRooms   = @()
    RoomNames       = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
    # No default - a per-session password typed into the Review step's
    # PasswordEntryPanel is required before every run (see btnCreate's
    # Add_Click validation). A hardcoded fallback here would mean every
    # tenant this tool has ever touched got the same predictable password
    # unless the admin remembered to change it.
    Password      = $null
}
$ui.lstRoomNames.ItemsSource = $Script:State.RoomNames

# A hashtable (reference type), not a plain $Script:-scoped bool: the
# Cancel button's click handler and the retry loops' CancelCheck
# scriptblocks need to observe the SAME flag changing over time, and a
# CancelCheck scriptblock is built (and .GetNewClosure()'d) from inside
# the already-closed-over Create/Apply handler - the same nested-closure
# situation that broke a plain $Script:State.Password reference earlier.
# Capturing this hashtable into a plain local first, then mutating its
# *contents* rather than reassigning the variable, sidesteps both
# problems at once (see $password in the Create/Apply handler for the
# same pattern).
$Script:CancelState = @{ Requested = $false }

function ConvertTo-SafeLocalPart([string]$Text) {
    $safe = ($Text -replace '[^a-zA-Z0-9\-\.]', '')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = [guid]::NewGuid().ToString('N').Substring(0, 8) }
    return $safe
}

# Set-Place's -CountryOrRegion wants a 2-letter ISO country code (e.g.
# "SE" for Sweden), not a country name - a plain text field asking for
# that was an easy way to get it wrong with no feedback. Built from .NET's
# own region data instead of a hand-typed list, so it's complete and needs
# no maintenance. The blank first entry keeps the existing "blank field =
# leave out of Set-Place" behavior (Set-RoomPlaceInfo already treats an
# empty CountryOrRegion as "not provided").
$CountryList = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::SpecificCultures) |
    ForEach-Object { try { [System.Globalization.RegionInfo]::new($_.Name) } catch { $null } } |
    Where-Object { $_ } |
    Group-Object TwoLetterISORegionName |
    ForEach-Object { [pscustomobject]@{ Name = $_.Group[0].EnglishName; Code = $_.Name } } |
    Sort-Object Name
$ui.cmbCountry.ItemsSource = @([pscustomobject]@{ Name = '(leave blank / unchanged)'; Code = '' }) + $CountryList
$ui.cmbCountry.SelectedIndex = 0

#========================================================#
# Step navigation
#========================================================#
function Get-RoomListSummaryText {
    if ($ui.radSkipRoomList.IsChecked) { return "(not changing Room List membership)" }
    if ($ui.radUseExistingRoomList.IsChecked) {
        if ($ui.lstRoomLists.SelectedItem) { return $ui.lstRoomLists.SelectedItem.Name } else { return '(none selected)' }
    }
    $previewName = ConvertTo-SafeLocalPart $ui.txtNewRoomListName.Text
    return "$($ui.txtNewRoomListName.Text) (new - $previewName@$($Script:State.DefaultDomain))"
}

function Get-CalendarSummaryText {
    if ($ui.radSkipCalendar.IsChecked) { return '(not changing calendar processing)' }
    if ($ui.radStandardMode.IsChecked) { return 'Standard (default settings)' }
    return 'Custom'
}

function Update-ReviewSummary {
    if ($Script:State.Mode -eq 'Create') {
        $ui.txtReviewHeading.Text = 'Review & create'
        $ui.txtReviewSub.Text = 'Everything below will be applied to each room you added.'
        $ui.btnCreate.Content = 'Create Rooms'
        $ui.chkResetPassword.Visibility = 'Collapsed'
        $ui.PasswordEntryPanel.Visibility = 'Visible'

        $caGroupDesc = if ($ui.radUseExistingCAGroup.IsChecked) {
            if ($ui.lstCAGroups.SelectedItem) { $ui.lstCAGroups.SelectedItem.DisplayName } else { '(none selected)' }
        } else {
            "$($ui.txtNewCAGroupName.Text) (new)"
        }
        $domain = if ($ui.cmbDomain.SelectedItem) { $ui.cmbDomain.SelectedItem } else { '(none selected)' }

        $ui.txtReviewSummary.Text = @"
Rooms to create ($($Script:State.RoomNames.Count)): $($Script:State.RoomNames -join ', ')
Domain: $domain
Room List: $(Get-RoomListSummaryText)
Conditional Access exclusion group: $caGroupDesc
Calendar processing: $(Get-CalendarSummaryText)
Password: set below (same for every room in this run)
"@
    } else {
        $ui.txtReviewHeading.Text = 'Review & apply'
        $ui.txtReviewSub.Text = 'Everything below will be applied to each room you selected. Anything not checked/changed below is left exactly as it is.'
        $ui.btnCreate.Content = 'Apply Changes'
        $ui.chkResetPassword.Visibility = 'Visible'
        $ui.PasswordEntryPanel.Visibility = if ($ui.chkResetPassword.IsChecked) { 'Visible' } else { 'Collapsed' }

        $selectedNames = @($ui.lstExistingRooms.SelectedItems) | ForEach-Object { $_.DisplayName }
        $passwordLine = if ($ui.chkResetPassword.IsChecked) { 'Password will be reset (set it below).' } else { 'Password: not changed' }
        $caGroupDesc = if ($ui.radUseExistingCAGroup.IsChecked) {
            if ($ui.lstCAGroups.SelectedItem) { $ui.lstCAGroups.SelectedItem.DisplayName } else { '(none selected)' }
        } else {
            "$($ui.txtNewCAGroupName.Text) (new)"
        }

        $ui.txtReviewSummary.Text = @"
Rooms to edit ($($selectedNames.Count)): $($selectedNames -join ', ')
Room List: $(Get-RoomListSummaryText)
Conditional Access exclusion group: $caGroupDesc (rooms already in it are left alone)
Calendar processing: $(Get-CalendarSummaryText)
$passwordLine
"@
    }
}

function Show-Step([int]$Step) {
    for ($i = 0; $i -lt $StepPanels.Count; $i++) {
        $isActive = ($i + 1) -eq $Step
        $StepPanels[$i].Visibility = if ($isActive) { 'Visible' } else { 'Collapsed' }
        $StepLabels[$i].Foreground = if ($isActive) { Get-Brush '#FFFFFF' } else { Get-Brush '#93979E' }
    }
    $Script:State.CurrentStep = $Step
    $ui.btnBack.IsEnabled = ($Step -gt 1)
    # On the last step, "Next" turns into "Exit" in the same slot rather
    # than disappearing - its click handler below checks CurrentStep and
    # confirms before actually closing, so a leftover habit of clicking
    # "Next" one more time doesn't close the app unintentionally.
    $ui.btnNext.Content = if ($Step -eq $TotalSteps) { 'Exit' } else { 'Next' }
    $ui.btnNext.Visibility = 'Visible'
    $ui.txtGlobalError.Text = ''
    if ($Step -eq $TotalSteps) { Update-ReviewSummary }
}

function Test-StepValid([int]$Step) {
    switch ($Step) {
        1 {
            if (-not $Script:State.Connected) {
                $ui.txtGlobalError.Text = 'Connect first.'
                return $false
            }
        }
        3 {
            if ($ui.radSkipRoomList.IsChecked) {
                # nothing to validate - not changing Room List membership
            } elseif ($ui.radUseExistingRoomList.IsChecked) {
                if (-not $ui.lstRoomLists.SelectedItem) {
                    $ui.txtGlobalError.Text = 'Select a Room List, or switch to "Create a new Room List".'
                    return $false
                }
            } elseif ([string]::IsNullOrWhiteSpace($ui.txtNewRoomListName.Text)) {
                $ui.txtGlobalError.Text = 'Enter a name for the new Room List.'
                return $false
            }
        }
        4 {
            if ($ui.radUseExistingCAGroup.IsChecked) {
                if (-not $ui.lstCAGroups.SelectedItem) {
                    $ui.txtGlobalError.Text = 'Select a group, or switch to "Create a new group".'
                    return $false
                }
            } elseif ([string]::IsNullOrWhiteSpace($ui.txtNewCAGroupName.Text)) {
                $ui.txtGlobalError.Text = 'Enter a name for the new group.'
                return $false
            }
        }
        6 {
            if ($Script:State.Mode -eq 'Create') {
                if ($Script:State.RoomNames.Count -eq 0) {
                    $ui.txtGlobalError.Text = 'Add at least one room name.'
                    return $false
                }
                if (-not $ui.cmbDomain.SelectedItem) {
                    $ui.txtGlobalError.Text = 'Select a domain.'
                    return $false
                }
            } else {
                if ($ui.lstExistingRooms.SelectedItems.Count -eq 0) {
                    $ui.txtGlobalError.Text = 'Select at least one existing room to edit.'
                    return $false
                }
            }
        }
    }
    return $true
}

function Get-AdjacentVisibleStep([int]$FromStep, [int]$Direction) {
    <#
        Walks Next (+1) or Back (-1) from $FromStep, skipping the Security
        Group and SSPR Exclusion steps entirely while in Edit mode - both
        only apply to newly-created rooms.
    #>
    $step = $FromStep + $Direction
    while ($Script:State.Mode -eq 'Edit' -and $CreateOnlySteps -contains $step) { $step += $Direction }
    return $step
}

$ui.btnNext.Add_Click({
    if ($Script:State.CurrentStep -eq $TotalSteps) {
        $confirm = [System.Windows.MessageBox]::Show(
            'Are you sure you want to exit?',
            'Exit Meeting Room Provisioning?',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) { $Window.Close() }
        return
    }
    if (-not (Test-StepValid -Step $Script:State.CurrentStep)) { return }
    $nextStep = Get-AdjacentVisibleStep -FromStep $Script:State.CurrentStep -Direction 1
    if ($nextStep -le $TotalSteps) { Show-Step -Step $nextStep }
}.GetNewClosure())
$ui.btnBack.Add_Click({
    $prevStep = Get-AdjacentVisibleStep -FromStep $Script:State.CurrentStep -Direction -1
    if ($prevStep -ge 1) { Show-Step -Step $prevStep }
}.GetNewClosure())

#========================================================#
# Step 1: Connect
#========================================================#
$ui.btnConnect.Add_Click({
    $ui.btnConnect.IsEnabled = $false
    $ui.progConnect.Visibility = 'Visible'
    $ui.progConnect.Maximum = 1
    $ui.progConnect.Value = 0
    $ui.txtConnectStatus.Foreground = Get-Brush '#5B6069'
    $ui.txtConnectStatus.Text = 'Checking required modules (this can take a while the first time)...'
    Sync-UI
    try {
        # A hashtable (reference type), not a plain int, for the same reason
        # $Script:CancelState is one elsewhere in this file: $installProgress
        # below needs to keep updating the SAME tracker across repeated
        # calls, and a plain int captured by .GetNewClosure() would freeze
        # at whatever value it had when the closure was created.
        #
        # Install-RoomProvisioningModules reports one tick per actual
        # operation - each disconnect, each old module version removed,
        # each module installed, each module imported - rather than one
        # per coarse phase, so the bar's fill actually tracks how much work
        # is left instead of jumping in a few big, uneven steps. -ExtraSteps
        # 2 reserves room in that SAME total for the two phases this handler
        # drives itself afterward (sign-in, initial tenant queries), so the
        # bar doesn't jump backward the moment those start.
        $connectProgress = @{ Step = 0; Total = 1 }
        $installProgress = {
            param($msg, $step, $total)
            $connectProgress.Step = $step
            $connectProgress.Total = $total
            $ui.progConnect.Maximum = $total
            $ui.progConnect.Value = $step
            $ui.txtConnectStatus.Text = $msg
            Sync-UI
        }.GetNewClosure()
        Install-RoomProvisioningModules -ProgressCallback $installProgress -ExtraSteps 2 | Out-Null

        $connectProgress.Step++
        $ui.progConnect.Value = $connectProgress.Step
        $ui.txtConnectStatus.Text = 'Signing in - look for the Exchange Online sign-in popup, then the Microsoft Graph sign-in popup.'
        Sync-UI
        Connect-RoomProvisioningServices

        $connectProgress.Step++
        $ui.progConnect.Value = $connectProgress.Step
        $ui.txtConnectStatus.Text = 'Connected. Loading Room Lists, groups and domains from your tenant...'
        Sync-UI
        $Script:State.RoomLists = @(Get-ExistingRoomLists)
        $ui.lstRoomLists.ItemsSource = $Script:State.RoomLists

        $Script:State.CAGroups = @(Get-ConditionalAccessExcludedGroups)
        $ui.lstCAGroups.ItemsSource = $Script:State.CAGroups

        # Graph only exposes whether SSPR is on tenant-wide - there is no
        # API to read or set which group(s) it's scoped to (confirmed by
        # testing the stable/beta SDKs and raw REST calls). See README
        # "SSPR exclusion" for what that limitation means for this step.
        $Script:State.SsprEnabled = [bool](Test-SelfServicePasswordResetEnabled)
        if ($Script:State.SsprEnabled) {
            $ui.txtSsprStatus.Text = 'Self-Service Password Reset is enabled in this tenant.'
            $ui.SsprDisabledPanel.Visibility = 'Collapsed'
            $Script:State.SsprGroupExists = [bool](Get-SsprExclusionGroup)
            $ui.SsprGroupFoundPanel.Visibility = if ($Script:State.SsprGroupExists) { 'Visible' } else { 'Collapsed' }
            $ui.SsprGroupMissingPanel.Visibility = if ($Script:State.SsprGroupExists) { 'Collapsed' } else { 'Visible' }
        } else {
            $ui.txtSsprStatus.Text = 'Self-Service Password Reset is not enabled in this tenant.'
            $ui.SsprDisabledPanel.Visibility = 'Visible'
            $ui.SsprGroupFoundPanel.Visibility = 'Collapsed'
            $ui.SsprGroupMissingPanel.Visibility = 'Collapsed'
        }

        $Script:State.Domains = @(Get-TenantDomains)
        $ui.cmbDomain.ItemsSource = $Script:State.Domains
        if ($Script:State.Domains.Count -gt 0) { $ui.cmbDomain.SelectedIndex = 0 }
        $Script:State.DefaultDomain = Get-DefaultTenantDomain
        if (-not $Script:State.DefaultDomain -and $Script:State.Domains.Count -gt 0) { $Script:State.DefaultDomain = $Script:State.Domains[0] }

        $ui.txtConnectStatus.Text = 'Checking existing room mailboxes...'
        Sync-UI
        $existingRooms = @(Get-ExistingRoomMailboxes)
        $Script:State.ExistingRooms = $existingRooms
        $ui.lstExistingRooms.ItemsSource = $existingRooms

        $licensedLines = [System.Collections.Generic.List[string]]::new()
        foreach ($room in $existingRooms) {
            $info = Get-RoomLicenseInfo -UserPrincipalName $room.UserPrincipalName
            if ($info.HasLicense) { $licensedLines.Add("- $($room.DisplayName): $($info.Licenses -join ', ')") }
        }
        if ($existingRooms.Count -eq 0) {
            $ui.txtLicenseInfo.Text = 'No existing room mailboxes were found in this tenant.'
        } elseif ($licensedLines.Count -gt 0) {
            $ui.txtLicenseInfo.Text = "$($licensedLines.Count) of $($existingRooms.Count) existing room mailbox(es) already have a license assigned:`n" +
                ($licensedLines -join "`n") +
                "`n`nCheck whether the new rooms you're creating will need one too. This tool does not assign licenses - do that in the Microsoft 365 admin center once purchased."
        } else {
            $ui.txtLicenseInfo.Text = "None of the $($existingRooms.Count) existing room mailbox(es) currently have a license assigned. If the new rooms need one, remember to purchase and assign it separately - this tool does not assign licenses."
        }
        $ui.LicenseInfoCard.Visibility = 'Visible'

        # Shows exactly which Graph delegated scopes actually got consented
        # for this sign-in, which account signed in, and - crucially -
        # that account's currently ACTIVE directory roles. That last part
        # matters specifically because a PIM-eligible (not activated)
        # Global Administrator assignment does not appear in a live
        # memberOf query and does not carry the role's permissions for
        # this token, even though the person genuinely holds the role and
        # a portal page might still show it. If a later step fails with
        # Authorization_RequestDenied despite "I'm a Global Admin", check
        # whether Global Administrator actually appears in this list.
        $graphContext = Get-MgContext
        $grantedScopes = if ($graphContext) { $graphContext.Scopes -join ', ' } else { '(none)' }
        $activeRoles = try {
            @(Get-MgUserMemberOf -UserId $graphContext.Account -All -ErrorAction Stop |
                Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' } |
                ForEach-Object { $_.AdditionalProperties['displayName'] }) -join ', '
        } catch { "(could not check: $($_.Exception.Message))" }
        if ([string]::IsNullOrWhiteSpace($activeRoles)) { $activeRoles = '(none active - if you expected to see Global Administrator or similar here, and this tenant uses PIM, the role is likely eligible but not activated for this sign-in)' }

        $Script:State.Connected = $true
        $ui.txtConnectStatus.Foreground = Get-Brush '#1F8A54'
        $ui.txtConnectStatus.Text = "Connected successfully as $($graphContext.Account). Click Next to continue.`n`nGranted Graph scopes: $grantedScopes`n`nActive directory roles for this sign-in: $activeRoles"
    } catch {
        $ui.txtConnectStatus.Foreground = Get-Brush '#C0392B'
        $ui.txtConnectStatus.Text = "Connection failed:`n$(Get-DiagnosticErrorText -ErrorRecord $_)"
        $ui.btnConnect.IsEnabled = $true
    } finally {
        $ui.progConnect.Visibility = 'Collapsed'
    }
}.GetNewClosure())

#========================================================#
# Mode: Create new rooms vs Edit existing rooms
#========================================================#
function Update-RoomListAddressPreview {
    if ($ui.radCreateNewRoomList.IsChecked -ne $true) {
        $ui.txtNewRoomListAddressPreview.Text = ''
        return
    }
    $previewName = ConvertTo-SafeLocalPart $ui.txtNewRoomListName.Text
    $domainPart = if ($Script:State.DefaultDomain) { $Script:State.DefaultDomain } else { '(connect first)' }
    $ui.txtNewRoomListAddressPreview.Text = "Will be created as: $previewName@$domainPart"
}

$ui.radModeCreate.Add_Checked({
    $Script:State.Mode = 'Create'
    $ui.radSkipRoomList.Visibility = 'Collapsed'
    if ($ui.radSkipRoomList.IsChecked) { $ui.radUseExistingRoomList.IsChecked = $true }
    $ui.radSkipCalendar.Visibility = 'Collapsed'
    if ($ui.radSkipCalendar.IsChecked) { $ui.radStandardMode.IsChecked = $true }
    $ui.CreateRoomNamingPanel.Visibility = 'Visible'
    $ui.EditRoomSelectionPanel.Visibility = 'Collapsed'
}.GetNewClosure())

$ui.radModeEdit.Add_Checked({
    $Script:State.Mode = 'Edit'
    $ui.radSkipRoomList.Visibility = 'Visible'
    $ui.radSkipRoomList.IsChecked = $true
    $ui.radSkipCalendar.Visibility = 'Visible'
    $ui.radSkipCalendar.IsChecked = $true
    $ui.CreateRoomNamingPanel.Visibility = 'Collapsed'
    $ui.EditRoomSelectionPanel.Visibility = 'Visible'
}.GetNewClosure())

$ui.txtNewRoomListName.Add_TextChanged({ Update-RoomListAddressPreview }.GetNewClosure())
$ui.radCreateNewRoomList.Add_Checked({ Update-RoomListAddressPreview }.GetNewClosure())
$ui.radUseExistingRoomList.Add_Checked({ $ui.txtNewRoomListAddressPreview.Text = '' }.GetNewClosure())
$ui.radSkipRoomList.Add_Checked({ $ui.txtNewRoomListAddressPreview.Text = '' }.GetNewClosure())

#========================================================#
# Step 4: Room names
#========================================================#
$AddRoomNameAction = {
    $name = $ui.txtRoomNameInput.Text.Trim()
    if ($name -and -not $Script:State.RoomNames.Contains($name)) { $Script:State.RoomNames.Add($name) }
    $ui.txtRoomNameInput.Text = ''
    $ui.txtRoomNameInput.Focus() | Out-Null
}.GetNewClosure()
$ui.btnAddRoomName.Add_Click($AddRoomNameAction)
$ui.txtRoomNameInput.Add_KeyDown({
    param($senderObj, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return) { & $AddRoomNameAction }
}.GetNewClosure())
$ui.btnRemoveRoomName.Add_Click({
    $selected = $ui.lstRoomNames.SelectedItem
    if ($selected) { $Script:State.RoomNames.Remove($selected) | Out-Null }
}.GetNewClosure())

#========================================================#
# Step 6: Calendar processing custom mode toggles
#========================================================#
$ui.radCustomMode.Add_Checked({ $ui.CustomCalendarPanel.Visibility = 'Visible' }.GetNewClosure())
$ui.radStandardMode.Add_Checked({ $ui.CustomCalendarPanel.Visibility = 'Collapsed' }.GetNewClosure())
$ui.radSkipCalendar.Add_Checked({ $ui.CustomCalendarPanel.Visibility = 'Collapsed' }.GetNewClosure())
$ui.cmbBookingWindow.Add_SelectionChanged({
    $selected = $ui.cmbBookingWindow.SelectedItem
    $ui.txtBookingWindowCustomDays.Visibility = if ($selected -and $selected.Tag -eq 'custom') { 'Visible' } else { 'Collapsed' }
}.GetNewClosure())

# Edit mode only: the password entry panel on the Review step only makes
# sense when a reset is actually requested, and this checkbox lives on that
# same step - toggling it needs to refresh the panel's visibility (and the
# summary text) immediately, not just the next time the step is entered.
$ui.chkResetPassword.Add_Checked({ Update-ReviewSummary }.GetNewClosure())
$ui.chkResetPassword.Add_Unchecked({ Update-ReviewSummary }.GetNewClosure())

#========================================================#
# Cancel button - see the "no background thread" note on
# Invoke-WithRetryProgress for why a retry wait can even notice this
# click at all.
#========================================================#
$ui.btnCancel.Add_Click({
    $confirm = [System.Windows.MessageBox]::Show(
        "Are you sure you want to cancel?`n`nThe room currently being processed may be left partially configured - some of its settings applied, others not. Rooms already fully processed keep whatever was done to them.",
        'Cancel operation?',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
        $Script:CancelState.Requested = $true
        $ui.btnCancel.IsEnabled = $false
        $ui.txtProgressStatus.Text = 'Cancelling - finishing the current step, then stopping...'
        Sync-UI
    }
}.GetNewClosure())

#========================================================#
# Step 7: Create
#========================================================#
$ui.btnCreate.Add_Click({
    # Create mode always sets a password on new rooms; Edit mode only needs
    # one when the reset checkbox is checked. Validated here rather than in
    # Test-StepValid because this is the last step - Next turns into Exit on
    # it (see Show-Step), so Test-StepValid never runs for it.
    $passwordRequired = ($Script:State.Mode -eq 'Create') -or [bool]$ui.chkResetPassword.IsChecked
    if ($passwordRequired) {
        $enteredPassword = $ui.pwdRoomPassword.Password
        if ([string]::IsNullOrEmpty($enteredPassword)) {
            $ui.txtGlobalError.Text = 'Enter a password for the room(s).'
            return
        }
        if ($enteredPassword -ne $ui.pwdRoomPasswordConfirm.Password) {
            $ui.txtGlobalError.Text = 'Password and confirmation do not match.'
            return
        }
        $Script:State.Password = $enteredPassword
    }

    $ui.btnCreate.IsEnabled = $false
    $ui.btnCancel.IsEnabled = $true
    $ui.btnCancel.Visibility = 'Visible'
    $Script:CancelState.Requested = $false
    $ui.txtLog.Inlines.Clear()
    $ui.ResultCard.Visibility = 'Collapsed'
    $ui.txtGlobalError.Text = ''

    # $AddLog is a closed-over scriptblock, not a nested `function` - a
    # function defined inside this handler would NOT be visible once
    # invoked (as a callback) from inside a different module's function;
    # see the NOTE in RoomProvisioning.Connections.psm1. Every scriptblock
    # below that gets handed to a module function ends in .GetNewClosure()
    # for the same reason.
    #
    # Colors each line by what it says rather than tracking state
    # separately: failures/cancellation in red, recognizable completions
    # in green, everything else (section headers, "using existing X",
    # retry attempts) in cyan. txtLog uses Inlines (Run + LineBreak) from
    # here on instead of a plain .Text string, since a single TextBlock
    # can only have one color via .Text.
    $AddLog = {
        param([string]$msg)
        $color = if ($msg -match '(?i)FAILED|Cancelled') {
            '#F87171'
        } elseif ($msg -match '(?i)\b(created|configured|added|set after|installed|removed|processed|connected|applied|granted)\b') {
            '#86EFAC'
        } else {
            '#5EEAD4'
        }
        $run = New-Object System.Windows.Documents.Run($msg)
        $run.Foreground = Get-Brush $color
        $ui.txtLog.Inlines.Add($run)
        $ui.txtLog.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        $ui.LogScrollViewer.ScrollToEnd()
        Sync-UI
    }.GetNewClosure()

    try {
        # --- Room List (shared - "skip" only reachable in Edit mode) ---
        $roomListIdentity = $null
        if ($ui.radSkipRoomList.IsChecked) {
            & $AddLog 'Not changing Room List membership.'
        } elseif ($ui.radUseExistingRoomList.IsChecked) {
            $roomListIdentity = $ui.lstRoomLists.SelectedItem.Identity
            & $AddLog "Using existing Room List: $($ui.lstRoomLists.SelectedItem.Name)"
        } else {
            $newListLocalPart = ConvertTo-SafeLocalPart $ui.txtNewRoomListName.Text
            $newListAddress = "$newListLocalPart@$($Script:State.DefaultDomain)"
            & $AddLog "Creating Room List '$($ui.txtNewRoomListName.Text)' ($newListAddress)..."
            $newList = New-RoomList -Name $ui.txtNewRoomListName.Text -PrimarySmtpAddress $newListAddress
            $roomListIdentity = $newList.Identity
            & $AddLog 'Room List created.'
        }

        # --- Calendar processing (shared - "skip" only reachable in Edit mode) ---
        $calendarParams = $null
        if ($ui.radSkipCalendar.IsChecked) {
            & $AddLog 'Not changing calendar processing.'
        } elseif ($ui.radStandardMode.IsChecked) {
            $calendarParams = Get-StandardCalendarProcessingParams
        } else {
            $bookingWindowDays = if ($ui.cmbBookingWindow.SelectedItem.Tag -eq 'custom') {
                [int]$ui.txtBookingWindowCustomDays.Text
            } else {
                [int]$ui.cmbBookingWindow.SelectedItem.Tag
            }
            $answers = [pscustomobject]@{
                IsPrivate                 = [bool]$ui.chkIsPrivate.IsChecked
                AllowConflictingSeries    = [bool]$ui.chkAllowConflictingSeries.IsChecked
                ConflictPercentageAllowed = $ui.txtConflictPercentage.Text
                MaximumConflictInstances  = $ui.txtMaxConflictInstances.Text
                BookingWindowDays         = $bookingWindowDays
                RequireApproval           = [bool]$ui.chkRequireApproval.IsChecked
                ApprovalDelegates         = @($ui.txtApprovalDelegates.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                AllowRecurringMeetings    = [bool]$ui.chkAllowRecurring.IsChecked
                RemoveAttachments         = [bool]$ui.chkRemoveAttachments.IsChecked
                AllowExternalRequests     = [bool]$ui.chkAllowExternalRequests.IsChecked
                RemovePrivateFlag         = [bool]$ui.chkRemovePrivateFlag.IsChecked
                AdditionalResponseText    = $ui.txtAdditionalResponseText.Text
            }
            $calendarParams = ConvertTo-CalendarProcessingParams -Answers $answers
        }

        # --- Place info (shared - blank fields are left out, which for an
        # existing room simply means "leave that value as it already is") ---
        $capacityValue = 0
        [void][int]::TryParse($ui.txtCapacity.Text, [ref]$capacityValue)
        $placeInfo = @{
            Building        = $ui.txtBuilding.Text
            Capacity        = $capacityValue
            City            = $ui.txtCity.Text
            PostalCode      = $ui.txtPostalCode.Text
            State           = $ui.txtState.Text
            Street          = $ui.txtStreet.Text
            CountryOrRegion = $ui.cmbCountry.SelectedValue
        }

        $ui.ProgressPanel.Visibility = 'Visible'

        # Captured into a plain local here, rather than referencing
        # $Script:State.Password directly inside $pwAction below: a
        # scriptblock that calls .GetNewClosure() while it's already
        # running *inside* another closure (this whole Add_Click handler
        # is one) doesn't reliably capture explicitly scope-qualified
        # variables like $Script:X - confirmed by reproducing it in
        # isolation. Plain local variables close over correctly even
        # nested two levels deep, which is why $email/$caGroupId work
        # fine in the retry actions below but $Script:State.Password
        # didn't.
        $password = $Script:State.Password

        # Same reasoning as $password above, applied to the shared cancel
        # flag: $cancelState is a plain local holding a *reference* to the
        # same hashtable $Script:CancelState points at, so mutations the
        # Cancel button's handler makes to that hashtable's contents are
        # still visible through $cancelState here - but $cancelCheck's own
        # .GetNewClosure() only needs to close over the plain local, never
        # the $Script:-qualified name directly.
        $cancelState = $Script:CancelState
        $cancelCheck = { $cancelState.Requested }.GetNewClosure()
        $sleepStep = { Sync-UI }.GetNewClosure()

        if ($Script:State.Mode -eq 'Create') {
            # --- Conditional Access exclusion group (Create mode only) ---
            if ($ui.radUseExistingCAGroup.IsChecked) {
                $caGroupId = $ui.lstCAGroups.SelectedItem.Id
                $caGroupName = $ui.lstCAGroups.SelectedItem.DisplayName
                & $AddLog "Using existing CA-excluded group: $caGroupName"
                # Re-checked on every run, not just when the group is first
                # created: a CA policy added since this group was last used
                # wouldn't otherwise get the exclusion until someone
                # remembered to add it by hand.
                Sync-GroupExclusionAcrossConditionalAccessPolicies -GroupId $caGroupId -GroupDisplayName $caGroupName -LogCallback $AddLog
            } else {
                & $AddLog "Creating group '$($ui.txtNewCAGroupName.Text)' and excluding it from every Conditional Access policy..."
                $newGroup = New-ConditionalAccessExclusionGroup -DisplayName $ui.txtNewCAGroupName.Text -LogCallback $AddLog
                $caGroupId = $newGroup.Id
            }

            # --- SSPR exclusion group (only if SSPR is enabled - same lookup-or-create logic in the Edit branch below) ---
            # $ssprGroup is looked up (and, if just created, its MembershipRule
            # populated) ONCE here and then tracked locally as each room below
            # appends its own exclusion clause - re-reading it from Graph
            # between rooms would risk landing on a replica that hasn't caught
            # up with the previous room's write yet and clobbering it.
            $ssprGroup = $null
            if ($Script:State.SsprEnabled) {
                $ssprGroup = Get-SsprExclusionGroup
                if (-not $ssprGroup -and $ui.chkCreateSsprGroup.IsChecked) {
                    & $AddLog "Creating '$(Get-SsprGroupDisplayName)' dynamic group for SSPR exclusion..."
                    try {
                        $ssprGroup = New-SsprDynamicExclusionGroup -LogCallback $AddLog
                        & $AddLog "IMPORTANT: Graph cannot retarget SSPR itself - go to Entra admin center > Password reset > Properties and set the scope to this group by hand."
                    } catch {
                        & $AddLog "FAILED to create SSPR exclusion group: $($_.Exception.Message)"
                    }
                } elseif (-not $ssprGroup) {
                    & $AddLog "'$(Get-SsprGroupDisplayName)' group not found - skipping SSPR exclusion for these rooms."
                }
            }

            $domain = $ui.cmbDomain.SelectedItem
            $passwordFailedRooms = [System.Collections.Generic.List[string]]::new()
            $passwordSuccessCount = 0

            foreach ($roomName in @($Script:State.RoomNames)) {
                if ($cancelState.Requested) {
                    & $AddLog "Cancelled before processing $roomName - stopping here."
                    break
                }
                & $AddLog "=== $roomName ==="
                $localPart = ConvertTo-SafeLocalPart $roomName
                $email = "$localPart@$domain"

                $created = New-RoomMailboxIfMissing -EmailAddress $email -Password $password -Name $roomName
                if ($created.Error) {
                    & $AddLog "FAILED to create mailbox for $roomName`: $($created.Error)"
                    continue
                }
                & $AddLog $(if ($created.Created) { "Mailbox created ($email)." } else { "Mailbox already existed ($email)." })

                if ($roomListIdentity) {
                    try {
                        Add-RoomToRoomList -RoomListIdentity $roomListIdentity -RoomEmailAddress $email
                        & $AddLog 'Added to Room List.'
                    } catch { & $AddLog "FAILED to add to Room List: $($_.Exception.Message)" }
                }

                if ($calendarParams) {
                    try {
                        Set-RoomCalendarProcessing -Identity $email -CalendarParams $calendarParams
                        & $AddLog 'Calendar processing configured.'
                    } catch { & $AddLog "FAILED to configure calendar processing: $($_.Exception.Message)" }
                }

                $ui.txtProgressStatus.Text = "Setting place information for $roomName (waiting for directory replication)..."
                Sync-UI
                $placeAction = { Set-RoomPlaceInfo -Identity $email -PlaceInfo $placeInfo }.GetNewClosure()
                $placeProgress = {
                    param($attempt, $max)
                    $ui.progRetry.Maximum = $max
                    $ui.progRetry.Value = $attempt
                    $ui.txtProgressStatus.Text = "Setting place information for $roomName... attempt $attempt of $max"
                    Sync-UI
                }.GetNewClosure()
                $placeResult = Invoke-WithRetryProgress -Action $placeAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $placeProgress -LogCallback $AddLog -CancelCheck $cancelCheck -SleepStep $sleepStep
                if ($placeResult.Success) { $ui.progRetry.Value = $ui.progRetry.Maximum }
                & $AddLog $(if ($placeResult.Cancelled) { 'Cancelled while setting place information.' } elseif ($placeResult.Success) { "Place information set after $($placeResult.Attempts) attempt(s)." } else { "FAILED to set place information after $($placeResult.Attempts) attempts: $($placeResult.Error)" })
                if ($placeResult.Cancelled) { break }

                $ui.txtProgressStatus.Text = "Adding $roomName to the security group (waiting for directory replication)..."
                Sync-UI
                $groupAction = { Add-RoomToGroup -GroupId $caGroupId -UserPrincipalName $email }.GetNewClosure()
                $groupProgress = {
                    param($attempt, $max)
                    $ui.progRetry.Maximum = $max
                    $ui.progRetry.Value = $attempt
                    $ui.txtProgressStatus.Text = "Adding $roomName to the security group... attempt $attempt of $max"
                    Sync-UI
                }.GetNewClosure()
                $groupResult = Invoke-WithRetryProgress -Action $groupAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $groupProgress -LogCallback $AddLog -CancelCheck $cancelCheck -SleepStep $sleepStep
                if ($groupResult.Success) { $ui.progRetry.Value = $ui.progRetry.Maximum }
                & $AddLog $(if ($groupResult.Cancelled) { 'Cancelled while adding to the security group.' } elseif ($groupResult.Success) { "Added to security group after $($groupResult.Attempts) attempt(s)." } else { "FAILED to add to security group after $($groupResult.Attempts) attempts: $($groupResult.Error)" })
                if ($groupResult.Cancelled) { break }

                $ui.txtProgressStatus.Text = "Setting password for $roomName..."
                Sync-UI
                $pwAction = { Set-RoomPassword -UserPrincipalName $email -Password $password }.GetNewClosure()
                $pwProgress = {
                    param($attempt, $max)
                    $ui.progRetry.Maximum = $max
                    $ui.progRetry.Value = $attempt
                    $ui.txtProgressStatus.Text = "Setting password for $roomName... attempt $attempt of $max"
                    Sync-UI
                }.GetNewClosure()
                $pwResult = Invoke-WithRetryProgress -Action $pwAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $pwProgress -LogCallback $AddLog -CancelCheck $cancelCheck -SleepStep $sleepStep
                if ($pwResult.Success) { $ui.progRetry.Value = $ui.progRetry.Maximum }
                if ($pwResult.Cancelled) {
                    & $AddLog 'Cancelled while setting password.'
                    break
                } elseif ($pwResult.Success) {
                    $passwordSuccessCount++
                    & $AddLog "Password set after $($pwResult.Attempts) attempt(s)."
                } else {
                    $passwordFailedRooms.Add($roomName)
                    & $AddLog "FAILED to set password after $($pwResult.Attempts) attempts: $($pwResult.Error)"
                }

                if ($Script:State.SsprEnabled -and $ssprGroup) {
                    if ($ssprGroup.MembershipRule -like "*userPrincipalName -ne `"$email`"*") {
                        & $AddLog "Already excluded from '$(Get-SsprGroupDisplayName)'."
                    } else {
                        $newSsprRule = "($($ssprGroup.MembershipRule)) and (user.userPrincipalName -ne `"$email`")"
                        $ui.txtProgressStatus.Text = "Excluding $roomName from SSPR scope..."
                        Sync-UI
                        $ssprAction = { Add-RoomToSsprExclusionRule -GroupId $ssprGroup.Id -NewRule $newSsprRule }.GetNewClosure()
                        $ssprProgress = {
                            param($attempt, $max)
                            $ui.progRetry.Maximum = $max
                            $ui.progRetry.Value = $attempt
                            $ui.txtProgressStatus.Text = "Excluding $roomName from SSPR scope... attempt $attempt of $max"
                            Sync-UI
                        }.GetNewClosure()
                        $ssprResult = Invoke-WithRetryProgress -Action $ssprAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $ssprProgress -LogCallback $AddLog -CancelCheck $cancelCheck -SleepStep $sleepStep
                        if ($ssprResult.Success) { $ssprGroup.MembershipRule = $newSsprRule; $ui.progRetry.Value = $ui.progRetry.Maximum }
                        & $AddLog $(if ($ssprResult.Cancelled) { 'Cancelled while excluding from SSPR scope.' } elseif ($ssprResult.Success) { "Excluded from '$(Get-SsprGroupDisplayName)' after $($ssprResult.Attempts) attempt(s)." } else { "FAILED to exclude from '$(Get-SsprGroupDisplayName)' after $($ssprResult.Attempts) attempts: $($ssprResult.Error)" })
                        if ($ssprResult.Cancelled) { break }
                    }
                }
            }

            # Only claim the password actually took - a prior version of
            # this text always said "password set" regardless of whether
            # every attempt above had actually failed.
            if ($passwordFailedRooms.Count -eq 0) {
                $ui.txtResultPassword.Text = "Password for every room created above: $($Script:State.Password)"
            } elseif ($passwordSuccessCount -eq 0) {
                $ui.txtResultPassword.Text = "Password was NOT set for any room - see the run log above for the error. The room(s) do not have the intended password."
            } else {
                $ui.txtResultPassword.Text = "Password set for $passwordSuccessCount room(s). FAILED for: $($passwordFailedRooms -join ', ') - see the run log above."
            }
            $ui.txtResultLicenseReminder.Text = 'Reminder: no license was assigned automatically. If these rooms need one, purchase and assign it in the Microsoft 365 admin center.'
            $ui.txtResultSsprReminder.Text = if ($Script:State.SsprEnabled) { "Reminder: SSPR is enabled in this tenant. Graph has no API to read or set its group scope, so go check Password reset > Properties in the Entra admin center to confirm it's scoped correctly." } else { '' }
        } else {
            # --- Edit mode: existing rooms only, no mailbox creation, no security group step ---
            $resetPassword = [bool]$ui.chkResetPassword.IsChecked
            $passwordFailedRooms = [System.Collections.Generic.List[string]]::new()
            $passwordSuccessCount = 0

            # Same pick-existing-or-create-new logic as the Create branch
            # below - a room being edited may never have been added to the
            # CA exclusion group (e.g. it was created before this tool
            # managed that, or the group didn't exist yet), so Edit mode
            # gets the same choice instead of only ever touching rooms at
            # creation time. Add-RoomToGroup is idempotent (skips if the
            # room is already a member), so rooms already in the group are
            # simply left alone.
            if ($ui.radUseExistingCAGroup.IsChecked) {
                $caGroupId = $ui.lstCAGroups.SelectedItem.Id
                $caGroupName = $ui.lstCAGroups.SelectedItem.DisplayName
                & $AddLog "Using existing CA-excluded group: $caGroupName"
                Sync-GroupExclusionAcrossConditionalAccessPolicies -GroupId $caGroupId -GroupDisplayName $caGroupName -LogCallback $AddLog
            } else {
                & $AddLog "Creating group '$($ui.txtNewCAGroupName.Text)' and excluding it from every Conditional Access policy..."
                $newGroup = New-ConditionalAccessExclusionGroup -DisplayName $ui.txtNewCAGroupName.Text -LogCallback $AddLog
                $caGroupId = $newGroup.Id
            }

            # Same lookup-or-create logic as the Create branch above - Edit
            # mode can also create the "SSPR Users" group if it's missing,
            # since editing existing rooms is exactly how a room from before
            # this feature existed gets excluded.
            $ssprGroup = $null
            if ($Script:State.SsprEnabled) {
                $ssprGroup = Get-SsprExclusionGroup
                if (-not $ssprGroup -and $ui.chkCreateSsprGroup.IsChecked) {
                    & $AddLog "Creating '$(Get-SsprGroupDisplayName)' dynamic group for SSPR exclusion..."
                    try {
                        $ssprGroup = New-SsprDynamicExclusionGroup -LogCallback $AddLog
                        & $AddLog "IMPORTANT: Graph cannot retarget SSPR itself - go to Entra admin center > Password reset > Properties and set the scope to this group by hand."
                    } catch {
                        & $AddLog "FAILED to create SSPR exclusion group: $($_.Exception.Message)"
                    }
                } elseif (-not $ssprGroup) {
                    & $AddLog "'$(Get-SsprGroupDisplayName)' group not found - skipping SSPR exclusion for these rooms."
                }
            }

            foreach ($room in @($ui.lstExistingRooms.SelectedItems)) {
                if ($cancelState.Requested) {
                    & $AddLog "Cancelled before processing $($room.DisplayName) - stopping here."
                    break
                }
                $email = $room.UserPrincipalName
                & $AddLog "=== $($room.DisplayName) ($email) ==="

                if ($roomListIdentity) {
                    try {
                        Add-RoomToRoomList -RoomListIdentity $roomListIdentity -RoomEmailAddress $email
                        & $AddLog 'Added to Room List.'
                    } catch { & $AddLog "FAILED to add to Room List: $($_.Exception.Message)" }
                }

                if ($calendarParams) {
                    try {
                        Set-RoomCalendarProcessing -Identity $email -CalendarParams $calendarParams
                        & $AddLog 'Calendar processing configured.'
                    } catch { & $AddLog "FAILED to configure calendar processing: $($_.Exception.Message)" }
                }

                $ui.txtProgressStatus.Text = "Setting place information for $($room.DisplayName)..."
                Sync-UI
                $placeAction = { Set-RoomPlaceInfo -Identity $email -PlaceInfo $placeInfo }.GetNewClosure()
                $placeProgress = {
                    param($attempt, $max)
                    $ui.progRetry.Maximum = $max
                    $ui.progRetry.Value = $attempt
                    $ui.txtProgressStatus.Text = "Setting place information for $($room.DisplayName)... attempt $attempt of $max"
                    Sync-UI
                }.GetNewClosure()
                $placeResult = Invoke-WithRetryProgress -Action $placeAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $placeProgress -LogCallback $AddLog -CancelCheck $cancelCheck -SleepStep $sleepStep
                if ($placeResult.Success) { $ui.progRetry.Value = $ui.progRetry.Maximum }
                & $AddLog $(if ($placeResult.Cancelled) { 'Cancelled while setting place information.' } elseif ($placeResult.Success) { "Place information set after $($placeResult.Attempts) attempt(s) (blank fields left unchanged)." } else { "FAILED to set place information after $($placeResult.Attempts) attempts: $($placeResult.Error)" })
                if ($placeResult.Cancelled) { break }

                $ui.txtProgressStatus.Text = "Adding $($room.DisplayName) to the security group (waiting for directory replication)..."
                Sync-UI
                $groupAction = { Add-RoomToGroup -GroupId $caGroupId -UserPrincipalName $email }.GetNewClosure()
                $groupProgress = {
                    param($attempt, $max)
                    $ui.progRetry.Maximum = $max
                    $ui.progRetry.Value = $attempt
                    $ui.txtProgressStatus.Text = "Adding $($room.DisplayName) to the security group... attempt $attempt of $max"
                    Sync-UI
                }.GetNewClosure()
                $groupResult = Invoke-WithRetryProgress -Action $groupAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $groupProgress -LogCallback $AddLog -CancelCheck $cancelCheck -SleepStep $sleepStep
                if ($groupResult.Success) { $ui.progRetry.Value = $ui.progRetry.Maximum }
                & $AddLog $(if ($groupResult.Cancelled) { 'Cancelled while adding to the security group.' } elseif ($groupResult.Success) { "Added to security group after $($groupResult.Attempts) attempt(s)." } else { "FAILED to add to security group after $($groupResult.Attempts) attempts: $($groupResult.Error)" })
                if ($groupResult.Cancelled) { break }

                if ($resetPassword) {
                    $ui.txtProgressStatus.Text = "Resetting password for $($room.DisplayName)..."
                    Sync-UI
                    $pwAction = { Set-RoomPassword -UserPrincipalName $email -Password $password }.GetNewClosure()
                    $pwProgress = {
                        param($attempt, $max)
                        $ui.progRetry.Maximum = $max
                        $ui.progRetry.Value = $attempt
                        $ui.txtProgressStatus.Text = "Resetting password for $($room.DisplayName)... attempt $attempt of $max"
                        Sync-UI
                    }.GetNewClosure()
                    $pwResult = Invoke-WithRetryProgress -Action $pwAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $pwProgress -LogCallback $AddLog -CancelCheck $cancelCheck -SleepStep $sleepStep
                    if ($pwResult.Success) { $ui.progRetry.Value = $ui.progRetry.Maximum }
                    if ($pwResult.Cancelled) {
                        & $AddLog 'Cancelled while resetting password.'
                        break
                    } elseif ($pwResult.Success) {
                        $passwordSuccessCount++
                        & $AddLog "Password reset after $($pwResult.Attempts) attempt(s)."
                    } else {
                        $passwordFailedRooms.Add($room.DisplayName)
                        & $AddLog "FAILED to reset password after $($pwResult.Attempts) attempts: $($pwResult.Error)"
                    }
                }

                # Not gated on $resetPassword: Edit mode is exactly how an
                # existing room edited before this feature existed gets
                # backfilled into the exclusion rule, independent of whether
                # its password is also being touched this run.
                if ($Script:State.SsprEnabled -and $ssprGroup) {
                    if ($ssprGroup.MembershipRule -like "*userPrincipalName -ne `"$email`"*") {
                        & $AddLog "Already excluded from '$(Get-SsprGroupDisplayName)'."
                    } else {
                        $newSsprRule = "($($ssprGroup.MembershipRule)) and (user.userPrincipalName -ne `"$email`")"
                        $ui.txtProgressStatus.Text = "Excluding $($room.DisplayName) from SSPR scope..."
                        Sync-UI
                        $ssprAction = { Add-RoomToSsprExclusionRule -GroupId $ssprGroup.Id -NewRule $newSsprRule }.GetNewClosure()
                        $ssprProgress = {
                            param($attempt, $max)
                            $ui.progRetry.Maximum = $max
                            $ui.progRetry.Value = $attempt
                            $ui.txtProgressStatus.Text = "Excluding $($room.DisplayName) from SSPR scope... attempt $attempt of $max"
                            Sync-UI
                        }.GetNewClosure()
                        $ssprResult = Invoke-WithRetryProgress -Action $ssprAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $ssprProgress -LogCallback $AddLog -CancelCheck $cancelCheck -SleepStep $sleepStep
                        if ($ssprResult.Success) { $ssprGroup.MembershipRule = $newSsprRule; $ui.progRetry.Value = $ui.progRetry.Maximum }
                        & $AddLog $(if ($ssprResult.Cancelled) { 'Cancelled while excluding from SSPR scope.' } elseif ($ssprResult.Success) { "Excluded from '$(Get-SsprGroupDisplayName)' after $($ssprResult.Attempts) attempt(s)." } else { "FAILED to exclude from '$(Get-SsprGroupDisplayName)' after $($ssprResult.Attempts) attempts: $($ssprResult.Error)" })
                        if ($ssprResult.Cancelled) { break }
                    }
                }
            }

            # Only claim the password actually took - a prior version of
            # this text always said "new password" whenever the checkbox
            # was checked, regardless of whether every attempt above had
            # actually failed.
            if (-not $resetPassword) {
                $ui.txtResultPassword.Text = 'Password was not changed.'
            } elseif ($passwordFailedRooms.Count -eq 0) {
                $ui.txtResultPassword.Text = "New password for the rooms above: $($Script:State.Password)"
            } elseif ($passwordSuccessCount -eq 0) {
                $ui.txtResultPassword.Text = "Password reset FAILED for every room - see the run log above for the error. The room(s) still have their old password."
            } else {
                $ui.txtResultPassword.Text = "Password reset for $passwordSuccessCount room(s). FAILED for: $($passwordFailedRooms -join ', ') - see the run log above."
            }
            $ui.txtResultLicenseReminder.Text = ''
            $ui.txtResultSsprReminder.Text = if ($Script:State.SsprEnabled) { "Reminder: SSPR is enabled in this tenant. Graph has no API to read or set its group scope, so go check Password reset > Properties in the Entra admin center to confirm it's scoped correctly." } else { '' }
        }

        # Colors/heads the result card to match what actually happened -
        # covers cancellation and password failures (both tracked above);
        # other per-step errors stay visible in the run log either way.
        if ($cancelState.Requested) {
            $ui.txtProgressStatus.Text = 'Cancelled.'
            $ui.txtResultHeading.Text = 'Cancelled'
            $ui.ResultCard.Background = Get-Brush '#F8F1E7'
            $ui.ResultCard.BorderBrush = Get-Brush '#B9770B'
            $ui.txtResultHeading.Foreground = Get-Brush '#B9770B'
            & $AddLog 'Cancelled by user - rooms already fully processed before the cancellation keep their changes.'
        } elseif ($passwordFailedRooms.Count -gt 0) {
            $ui.txtProgressStatus.Text = 'Done, with errors.'
            $ui.txtResultHeading.Text = 'Completed with errors'
            $ui.ResultCard.Background = Get-Brush '#F8F1E7'
            $ui.ResultCard.BorderBrush = Get-Brush '#B9770B'
            $ui.txtResultHeading.Foreground = Get-Brush '#B9770B'
            & $AddLog 'All rooms processed - see above for password failures.'
        } else {
            $ui.txtProgressStatus.Text = 'Done.'
            $ui.txtResultHeading.Text = 'Done'
            $ui.ResultCard.Background = Get-Brush '#ECF6F1'
            $ui.ResultCard.BorderBrush = Get-Brush '#1F8A54'
            $ui.txtResultHeading.Foreground = Get-Brush '#1F8A54'
            & $AddLog 'All rooms processed.'
        }
        $ui.ResultCard.Visibility = 'Visible'
    } catch {
        & $AddLog "FATAL ERROR: $($_.Exception.Message)"
        $ui.txtGlobalError.Text = $_.Exception.Message
    } finally {
        $ui.btnCreate.IsEnabled = $true
        $ui.btnCancel.Visibility = 'Collapsed'

        # $Script:State.CAGroups/SsprGroupExists are otherwise only ever
        # populated once, at Connect - so a group created or found during
        # THIS run would still show as "not found"/missing from the list if
        # the admin clicks Back afterward to process another batch of rooms
        # in the same session, instead of closing and reopening the tool.
        try {
            $Script:State.CAGroups = @(Get-ConditionalAccessExcludedGroups)
            $ui.lstCAGroups.ItemsSource = $Script:State.CAGroups
        } catch { }
        if ($Script:State.SsprEnabled) {
            try {
                $Script:State.SsprGroupExists = [bool](Get-SsprExclusionGroup)
                $ui.SsprGroupFoundPanel.Visibility = if ($Script:State.SsprGroupExists) { 'Visible' } else { 'Collapsed' }
                $ui.SsprGroupMissingPanel.Visibility = if ($Script:State.SsprGroupExists) { 'Collapsed' } else { 'Visible' }
            } catch { }
        }
    }
}.GetNewClosure())

$Window.Add_Closing({ Disconnect-RoomProvisioningServices }.GetNewClosure())

if ($PSVersionTable.PSEdition -ne 'Core') {
    $ui.txtConnectStatus.Foreground = Get-Brush '#B9770B'
    $ui.txtConnectStatus.Text = 'Running under Windows PowerShell 5.1 - PowerShell 7 was not found on this machine, so Microsoft Graph sign-in may fail with a "GetTokenAsync ... lacks an implementation" error (a known incompatibility between the Graph SDK and Windows PowerShell 5.1). Installing PowerShell 7 (winget install Microsoft.PowerShell) and re-running this tool is the reliable fix.'
}

Show-Step -Step 1
[void]$Window.ShowDialog()
