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

$IsElevated = Test-IsAdministrator
$IsWindowsPowerShellDesktop = $PSVersionTable.PSEdition -ne 'Core'
$Pwsh = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
$NeedsRelaunch = (-not $IsElevated) -or ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') -or ($IsWindowsPowerShellDesktop -and $Pwsh)

if ($NeedsRelaunch) {
    $exe = if ($IsWindowsPowerShellDesktop -and $Pwsh) { $Pwsh.Source } else { (Get-Process -Id $PID).Path }
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
    'lblStep1', 'lblStep2', 'lblStep3', 'lblStep4', 'lblStep5', 'lblStep6', 'lblStep7',
    'Step1Panel', 'Step2Panel', 'Step3Panel', 'Step4Panel', 'Step5Panel', 'Step6Panel', 'Step7Panel',
    'btnConnect', 'txtConnectStatus', 'LicenseInfoCard', 'txtLicenseInfo',
    'radUseExistingRoomList', 'lstRoomLists', 'radCreateNewRoomList', 'txtNewRoomListName', 'txtNewRoomListAddress',
    'radUseExistingCAGroup', 'lstCAGroups', 'radCreateNewCAGroup', 'txtNewCAGroupName',
    'txtRoomNameInput', 'btnAddRoomName', 'lstRoomNames', 'btnRemoveRoomName', 'cmbDomain',
    'txtBuilding', 'txtCapacity', 'txtCity', 'txtPostalCode', 'txtState', 'txtStreet', 'txtCountry',
    'radStandardMode', 'radCustomMode', 'CustomCalendarPanel',
    'chkIsPrivate', 'chkAllowConflictingSeries', 'txtConflictPercentage', 'txtMaxConflictInstances',
    'cmbBookingWindow', 'txtBookingWindowCustomDays', 'chkRequireApproval', 'txtApprovalDelegates',
    'chkAllowRecurring', 'chkRemoveAttachments', 'chkAllowExternalRequests', 'chkRemovePrivateFlag', 'txtAdditionalResponseText',
    'txtReviewSummary', 'btnCreate', 'ProgressPanel', 'txtProgressStatus', 'progRetry',
    'txtLog', 'LogScrollViewer', 'ResultCard', 'txtResultPassword', 'txtResultLicenseReminder',
    'btnBack', 'btnNext', 'txtGlobalError'
)
$ui = @{}
foreach ($name in $ElementNames) { $ui[$name] = $Window.FindName($name) }

$StepPanels = @($ui.Step1Panel, $ui.Step2Panel, $ui.Step3Panel, $ui.Step4Panel, $ui.Step5Panel, $ui.Step6Panel, $ui.Step7Panel)
$StepLabels = @($ui.lblStep1, $ui.lblStep2, $ui.lblStep3, $ui.lblStep4, $ui.lblStep5, $ui.lblStep6, $ui.lblStep7)
$TotalSteps = $StepPanels.Count

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
    CurrentStep = 1
    Connected   = $false
    RoomLists   = @()
    CAGroups    = @()
    Domains     = @()
    RoomNames   = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
    Password    = 'REDACTED-ROTATE-THIS-PASSWORD'
}
$ui.lstRoomNames.ItemsSource = $Script:State.RoomNames

#========================================================#
# Step navigation
#========================================================#
function Update-ReviewSummary {
    $roomListDesc = if ($ui.radUseExistingRoomList.IsChecked) {
        if ($ui.lstRoomLists.SelectedItem) { $ui.lstRoomLists.SelectedItem.Name } else { '(none selected)' }
    } else {
        "$($ui.txtNewRoomListName.Text) (new)"
    }
    $caGroupDesc = if ($ui.radUseExistingCAGroup.IsChecked) {
        if ($ui.lstCAGroups.SelectedItem) { $ui.lstCAGroups.SelectedItem.DisplayName } else { '(none selected)' }
    } else {
        "$($ui.txtNewCAGroupName.Text) (new)"
    }
    $domain = if ($ui.cmbDomain.SelectedItem) { $ui.cmbDomain.SelectedItem } else { '(none selected)' }
    $calendarMode = if ($ui.radStandardMode.IsChecked) { 'Standard (default settings)' } else { 'Custom' }

    $ui.txtReviewSummary.Text = @"
Rooms to create ($($Script:State.RoomNames.Count)): $($Script:State.RoomNames -join ', ')
Domain: $domain
Room List: $roomListDesc
Conditional Access exclusion group: $caGroupDesc
Calendar processing: $calendarMode
Password (same for every room): $($Script:State.Password)
"@
}

function Show-Step([int]$Step) {
    for ($i = 0; $i -lt $StepPanels.Count; $i++) {
        $isActive = ($i + 1) -eq $Step
        $StepPanels[$i].Visibility = if ($isActive) { 'Visible' } else { 'Collapsed' }
        $StepLabels[$i].Foreground = if ($isActive) { Get-Brush '#FFFFFF' } else { Get-Brush '#8CA0C4' }
    }
    $Script:State.CurrentStep = $Step
    $ui.btnBack.IsEnabled = ($Step -gt 1)
    $ui.btnNext.Visibility = if ($Step -eq $TotalSteps) { 'Collapsed' } else { 'Visible' }
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
        2 {
            if ($ui.radUseExistingRoomList.IsChecked) {
                if (-not $ui.lstRoomLists.SelectedItem) {
                    $ui.txtGlobalError.Text = 'Select a Room List, or switch to "Create a new Room List".'
                    return $false
                }
            } elseif ([string]::IsNullOrWhiteSpace($ui.txtNewRoomListName.Text) -or [string]::IsNullOrWhiteSpace($ui.txtNewRoomListAddress.Text)) {
                $ui.txtGlobalError.Text = 'Enter both a name and an email address for the new Room List.'
                return $false
            }
        }
        3 {
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
        4 {
            if ($Script:State.RoomNames.Count -eq 0) {
                $ui.txtGlobalError.Text = 'Add at least one room name.'
                return $false
            }
            if (-not $ui.cmbDomain.SelectedItem) {
                $ui.txtGlobalError.Text = 'Select a domain.'
                return $false
            }
        }
    }
    return $true
}

$ui.btnNext.Add_Click({
    if (-not (Test-StepValid -Step $Script:State.CurrentStep)) { return }
    if ($Script:State.CurrentStep -lt $TotalSteps) { Show-Step -Step ($Script:State.CurrentStep + 1) }
}.GetNewClosure())
$ui.btnBack.Add_Click({
    if ($Script:State.CurrentStep -gt 1) { Show-Step -Step ($Script:State.CurrentStep - 1) }
}.GetNewClosure())

#========================================================#
# Step 1: Connect
#========================================================#
$ui.btnConnect.Add_Click({
    $ui.btnConnect.IsEnabled = $false
    $ui.txtConnectStatus.Foreground = Get-Brush '#697586'
    $ui.txtConnectStatus.Text = 'Checking required modules (this can take a while the first time)...'
    Sync-UI
    try {
        $installProgress = {
            param($msg) $ui.txtConnectStatus.Text = $msg; Sync-UI
        }.GetNewClosure()
        Install-RoomProvisioningModules -ProgressCallback $installProgress | Out-Null

        $ui.txtConnectStatus.Text = 'Signing in to Exchange Online (look for its sign-in popup), then Microsoft Graph - for Graph, switch to the console window that opened alongside this app: it will show a code and https://microsoft.com/devicelogin to finish signing in in your browser.'
        Sync-UI
        Connect-RoomProvisioningServices

        $ui.txtConnectStatus.Text = 'Connected. Loading Room Lists, groups and domains from your tenant...'
        Sync-UI
        $Script:State.RoomLists = @(Get-ExistingRoomLists)
        $ui.lstRoomLists.ItemsSource = $Script:State.RoomLists

        $Script:State.CAGroups = @(Get-ConditionalAccessExcludedGroups)
        $ui.lstCAGroups.ItemsSource = $Script:State.CAGroups

        $Script:State.Domains = @(Get-TenantDomains)
        $ui.cmbDomain.ItemsSource = $Script:State.Domains
        if ($Script:State.Domains.Count -gt 0) { $ui.cmbDomain.SelectedIndex = 0 }

        $ui.txtConnectStatus.Text = 'Checking existing room mailboxes for assigned licenses...'
        Sync-UI
        $existingRooms = @(Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue)
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

        $Script:State.Connected = $true
        $ui.txtConnectStatus.Foreground = Get-Brush '#1E8E5A'
        $ui.txtConnectStatus.Text = 'Connected successfully. Click Next to continue.'
    } catch {
        $ui.txtConnectStatus.Foreground = Get-Brush '#C0392B'
        $ui.txtConnectStatus.Text = "Connection failed:`n$(Get-DiagnosticErrorText -ErrorRecord $_)"
        $ui.btnConnect.IsEnabled = $true
    }
}.GetNewClosure())

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
$ui.cmbBookingWindow.Add_SelectionChanged({
    $selected = $ui.cmbBookingWindow.SelectedItem
    $ui.txtBookingWindowCustomDays.Visibility = if ($selected -and $selected.Tag -eq 'custom') { 'Visible' } else { 'Collapsed' }
}.GetNewClosure())

#========================================================#
# Step 7: Create
#========================================================#
$ui.btnCreate.Add_Click({
    $ui.btnCreate.IsEnabled = $false
    $ui.txtLog.Text = ''
    $ui.ResultCard.Visibility = 'Collapsed'
    $ui.txtGlobalError.Text = ''

    # $AddLog is a closed-over scriptblock, not a nested `function` - a
    # function defined inside this handler would NOT be visible once
    # invoked (as a callback) from inside a different module's function;
    # see the NOTE in RoomProvisioning.Connections.psm1. Every scriptblock
    # below that gets handed to a module function ends in .GetNewClosure()
    # for the same reason.
    $AddLog = {
        param([string]$msg)
        $ui.txtLog.Text += "$msg`n"
        $ui.LogScrollViewer.ScrollToEnd()
        Sync-UI
    }.GetNewClosure()

    try {
        # --- Room List ---
        if ($ui.radUseExistingRoomList.IsChecked) {
            $roomListIdentity = $ui.lstRoomLists.SelectedItem.Identity
            & $AddLog "Using existing Room List: $($ui.lstRoomLists.SelectedItem.Name)"
        } else {
            & $AddLog "Creating Room List '$($ui.txtNewRoomListName.Text)'..."
            $newList = New-RoomList -Name $ui.txtNewRoomListName.Text -PrimarySmtpAddress $ui.txtNewRoomListAddress.Text
            $roomListIdentity = $newList.Identity
            & $AddLog 'Room List created.'
        }

        # --- Conditional Access exclusion group ---
        if ($ui.radUseExistingCAGroup.IsChecked) {
            $caGroupId = $ui.lstCAGroups.SelectedItem.Id
            & $AddLog "Using existing CA-excluded group: $($ui.lstCAGroups.SelectedItem.DisplayName)"
        } else {
            & $AddLog "Creating group '$($ui.txtNewCAGroupName.Text)' and excluding it from every Conditional Access policy..."
            $newGroup = New-ConditionalAccessExclusionGroup -DisplayName $ui.txtNewCAGroupName.Text -LogCallback $AddLog
            $caGroupId = $newGroup.Id
        }

        # --- Calendar processing ---
        if ($ui.radStandardMode.IsChecked) {
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

        # --- Place info (blank fields are simply left out) ---
        $capacityValue = 0
        [void][int]::TryParse($ui.txtCapacity.Text, [ref]$capacityValue)
        $placeInfo = @{
            Building        = $ui.txtBuilding.Text
            Capacity        = $capacityValue
            City            = $ui.txtCity.Text
            PostalCode      = $ui.txtPostalCode.Text
            State           = $ui.txtState.Text
            Street          = $ui.txtStreet.Text
            CountryOrRegion = $ui.txtCountry.Text
        }

        $domain = $ui.cmbDomain.SelectedItem
        $ui.ProgressPanel.Visibility = 'Visible'

        foreach ($roomName in @($Script:State.RoomNames)) {
            & $AddLog "=== $roomName ==="

            $localPart = ($roomName -replace '[^a-zA-Z0-9\-\.]', '')
            if ([string]::IsNullOrWhiteSpace($localPart)) { $localPart = [guid]::NewGuid().ToString('N').Substring(0, 8) }
            $email = "$localPart@$domain"

            $created = New-RoomMailboxIfMissing -EmailAddress $email -Password $Script:State.Password -Name $roomName
            if ($created.Error) {
                & $AddLog "FAILED to create mailbox for $roomName`: $($created.Error)"
                continue
            }
            & $AddLog $(if ($created.Created) { "Mailbox created ($email)." } else { "Mailbox already existed ($email)." })

            try {
                Add-RoomToRoomList -RoomListIdentity $roomListIdentity -RoomEmailAddress $email
                & $AddLog 'Added to Room List.'
            } catch { & $AddLog "FAILED to add to Room List: $($_.Exception.Message)" }

            try {
                Set-RoomCalendarProcessing -Identity $email -CalendarParams $calendarParams
                & $AddLog 'Calendar processing configured.'
            } catch { & $AddLog "FAILED to configure calendar processing: $($_.Exception.Message)" }

            try {
                Set-RoomPlaceInfo -Identity $email -PlaceInfo $placeInfo
                & $AddLog 'Place information set.'
            } catch { & $AddLog "FAILED to set place information: $($_.Exception.Message)" }

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
            $groupResult = Invoke-WithRetryProgress -Action $groupAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $groupProgress -LogCallback $AddLog
            & $AddLog $(if ($groupResult.Success) { "Added to security group after $($groupResult.Attempts) attempt(s)." } else { "FAILED to add to security group after $($groupResult.Attempts) attempts: $($groupResult.Error)" })

            $ui.txtProgressStatus.Text = "Setting password for $roomName..."
            Sync-UI
            $pwAction = { Set-RoomPassword -UserPrincipalName $email -Password $Script:State.Password }.GetNewClosure()
            $pwProgress = {
                param($attempt, $max)
                $ui.progRetry.Maximum = $max
                $ui.progRetry.Value = $attempt
                $ui.txtProgressStatus.Text = "Setting password for $roomName... attempt $attempt of $max"
                Sync-UI
            }.GetNewClosure()
            $pwResult = Invoke-WithRetryProgress -Action $pwAction -MaxRetries 10 -DelaySeconds 20 -ProgressCallback $pwProgress -LogCallback $AddLog
            & $AddLog $(if ($pwResult.Success) { "Password set after $($pwResult.Attempts) attempt(s)." } else { "FAILED to set password after $($pwResult.Attempts) attempts: $($pwResult.Error)" })
        }

        $ui.txtProgressStatus.Text = 'Done.'
        $ui.txtResultPassword.Text = "Password for every room created above: $($Script:State.Password)"
        $ui.txtResultLicenseReminder.Text = 'Reminder: no license was assigned automatically. If these rooms need one, purchase and assign it in the Microsoft 365 admin center.'
        $ui.ResultCard.Visibility = 'Visible'
        & $AddLog 'All rooms processed.'
    } catch {
        & $AddLog "FATAL ERROR: $($_.Exception.Message)"
        $ui.txtGlobalError.Text = $_.Exception.Message
    } finally {
        $ui.btnCreate.IsEnabled = $true
    }
}.GetNewClosure())

$Window.Add_Closing({ Disconnect-RoomProvisioningServices }.GetNewClosure())

if ($PSVersionTable.PSEdition -ne 'Core') {
    $ui.txtConnectStatus.Foreground = Get-Brush '#B8860B'
    $ui.txtConnectStatus.Text = 'Running under Windows PowerShell 5.1 - PowerShell 7 was not found on this machine, so Microsoft Graph sign-in may fail with a "GetTokenAsync ... lacks an implementation" error (a known incompatibility between the Graph SDK and Windows PowerShell 5.1). Installing PowerShell 7 (winget install Microsoft.PowerShell) and re-running this tool is the reliable fix.'
}

Show-Step -Step 1
[void]$Window.ShowDialog()
