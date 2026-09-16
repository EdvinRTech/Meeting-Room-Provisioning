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

#========================================================#
# WPF requires an STA (single-threaded apartment) thread.
# Windows PowerShell (powershell.exe) defaults to STA already; PowerShell 7
# (pwsh.exe) defaults to MTA, so relaunch ourselves with -STA if needed.
#========================================================#
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-File', "`"$PSCommandPath`"") -Wait
    exit
}

$ErrorActionPreference = 'Stop'
$ScriptRoot = $PSScriptRoot

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.Common.psm1') -Force
Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.Connections.psm1') -Force
Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.Exchange.psm1') -Force
Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.Graph.psm1') -Force
Import-Module (Join-Path $ScriptRoot 'Modules\RoomProvisioning.CalendarLogic.psm1') -Force

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
})
$ui.btnBack.Add_Click({
    if ($Script:State.CurrentStep -gt 1) { Show-Step -Step ($Script:State.CurrentStep - 1) }
})

#========================================================#
# Step 1: Connect
#========================================================#
$ui.btnConnect.Add_Click({
    $ui.btnConnect.IsEnabled = $false
    $ui.txtConnectStatus.Foreground = Get-Brush '#697586'
    $ui.txtConnectStatus.Text = 'Checking required modules (this can take a while the first time)...'
    Sync-UI
    try {
        Install-RoomProvisioningModules -ProgressCallback {
            param($msg) $ui.txtConnectStatus.Text = $msg; Sync-UI
        } | Out-Null

        $ui.txtConnectStatus.Text = 'Signing in - look for the Exchange Online and Microsoft Graph sign-in prompts...'
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
        $ui.txtConnectStatus.Text = "Connection failed: $($_.Exception.Message)"
        $ui.btnConnect.IsEnabled = $true
    }
})

#========================================================#
# Step 4: Room names
#========================================================#
$AddRoomNameAction = {
    $name = $ui.txtRoomNameInput.Text.Trim()
    if ($name -and -not $Script:State.RoomNames.Contains($name)) { $Script:State.RoomNames.Add($name) }
    $ui.txtRoomNameInput.Text = ''
    $ui.txtRoomNameInput.Focus() | Out-Null
}
$ui.btnAddRoomName.Add_Click($AddRoomNameAction)
$ui.txtRoomNameInput.Add_KeyDown({
    param($senderObj, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return) { & $AddRoomNameAction }
})
$ui.btnRemoveRoomName.Add_Click({
    $selected = $ui.lstRoomNames.SelectedItem
    if ($selected) { $Script:State.RoomNames.Remove($selected) | Out-Null }
})

#========================================================#
# Step 6: Calendar processing custom mode toggles
#========================================================#
$ui.radCustomMode.Add_Checked({ $ui.CustomCalendarPanel.Visibility = 'Visible' })
$ui.radStandardMode.Add_Checked({ $ui.CustomCalendarPanel.Visibility = 'Collapsed' })
$ui.cmbBookingWindow.Add_SelectionChanged({
    $selected = $ui.cmbBookingWindow.SelectedItem
    $ui.txtBookingWindowCustomDays.Visibility = if ($selected -and $selected.Tag -eq 'custom') { 'Visible' } else { 'Collapsed' }
})

#========================================================#
# Step 7: Create
#========================================================#
$ui.btnCreate.Add_Click({
    $ui.btnCreate.IsEnabled = $false
    $ui.txtLog.Text = ''
    $ui.ResultCard.Visibility = 'Collapsed'
    $ui.txtGlobalError.Text = ''

    function Add-Log([string]$msg) {
        $ui.txtLog.Text += "$msg`n"
        $ui.LogScrollViewer.ScrollToEnd()
        Sync-UI
    }

    try {
        # --- Room List ---
        if ($ui.radUseExistingRoomList.IsChecked) {
            $roomListIdentity = $ui.lstRoomLists.SelectedItem.Identity
            Add-Log "Using existing Room List: $($ui.lstRoomLists.SelectedItem.Name)"
        } else {
            Add-Log "Creating Room List '$($ui.txtNewRoomListName.Text)'..."
            $newList = New-RoomList -Name $ui.txtNewRoomListName.Text -PrimarySmtpAddress $ui.txtNewRoomListAddress.Text
            $roomListIdentity = $newList.Identity
            Add-Log 'Room List created.'
        }

        # --- Conditional Access exclusion group ---
        if ($ui.radUseExistingCAGroup.IsChecked) {
            $caGroupId = $ui.lstCAGroups.SelectedItem.Id
            Add-Log "Using existing CA-excluded group: $($ui.lstCAGroups.SelectedItem.DisplayName)"
        } else {
            Add-Log "Creating group '$($ui.txtNewCAGroupName.Text)' and excluding it from every Conditional Access policy..."
            $newGroup = New-ConditionalAccessExclusionGroup -DisplayName $ui.txtNewCAGroupName.Text -LogCallback { param($m) Add-Log $m }
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
            Add-Log "=== $roomName ==="

            $localPart = ($roomName -replace '[^a-zA-Z0-9\-\.]', '')
            if ([string]::IsNullOrWhiteSpace($localPart)) { $localPart = [guid]::NewGuid().ToString('N').Substring(0, 8) }
            $email = "$localPart@$domain"

            $created = New-RoomMailboxIfMissing -EmailAddress $email -Password $Script:State.Password -Name $roomName
            if ($created.Error) {
                Add-Log "FAILED to create mailbox for $roomName`: $($created.Error)"
                continue
            }
            Add-Log $(if ($created.Created) { "Mailbox created ($email)." } else { "Mailbox already existed ($email)." })

            try {
                Add-RoomToRoomList -RoomListIdentity $roomListIdentity -RoomEmailAddress $email
                Add-Log 'Added to Room List.'
            } catch { Add-Log "FAILED to add to Room List: $($_.Exception.Message)" }

            try {
                Set-RoomCalendarProcessing -Identity $email -CalendarParams $calendarParams
                Add-Log 'Calendar processing configured.'
            } catch { Add-Log "FAILED to configure calendar processing: $($_.Exception.Message)" }

            try {
                Set-RoomPlaceInfo -Identity $email -PlaceInfo $placeInfo
                Add-Log 'Place information set.'
            } catch { Add-Log "FAILED to set place information: $($_.Exception.Message)" }

            $ui.txtProgressStatus.Text = "Adding $roomName to the security group (waiting for directory replication)..."
            Sync-UI
            $groupResult = Invoke-WithRetryProgress -Action { Add-RoomToGroup -GroupId $caGroupId -UserPrincipalName $email } `
                -MaxRetries 10 -DelaySeconds 20 `
                -ProgressCallback {
                    param($attempt, $max)
                    $ui.progRetry.Maximum = $max
                    $ui.progRetry.Value = $attempt
                    $ui.txtProgressStatus.Text = "Adding $roomName to the security group... attempt $attempt of $max"
                    Sync-UI
                } `
                -LogCallback { param($m) Add-Log $m }
            Add-Log $(if ($groupResult.Success) { "Added to security group after $($groupResult.Attempts) attempt(s)." } else { "FAILED to add to security group after $($groupResult.Attempts) attempts: $($groupResult.Error)" })

            $ui.txtProgressStatus.Text = "Setting password for $roomName..."
            Sync-UI
            $pwResult = Invoke-WithRetryProgress -Action { Set-RoomPassword -UserPrincipalName $email -Password $Script:State.Password } `
                -MaxRetries 10 -DelaySeconds 20 `
                -ProgressCallback {
                    param($attempt, $max)
                    $ui.progRetry.Maximum = $max
                    $ui.progRetry.Value = $attempt
                    $ui.txtProgressStatus.Text = "Setting password for $roomName... attempt $attempt of $max"
                    Sync-UI
                } `
                -LogCallback { param($m) Add-Log $m }
            Add-Log $(if ($pwResult.Success) { "Password set after $($pwResult.Attempts) attempt(s)." } else { "FAILED to set password after $($pwResult.Attempts) attempts: $($pwResult.Error)" })
        }

        $ui.txtProgressStatus.Text = 'Done.'
        $ui.txtResultPassword.Text = "Password for every room created above: $($Script:State.Password)"
        $ui.txtResultLicenseReminder.Text = 'Reminder: no license was assigned automatically. If these rooms need one, purchase and assign it in the Microsoft 365 admin center.'
        $ui.ResultCard.Visibility = 'Visible'
        Add-Log 'All rooms processed.'
    } catch {
        Add-Log "FATAL ERROR: $($_.Exception.Message)"
        $ui.txtGlobalError.Text = $_.Exception.Message
    } finally {
        $ui.btnCreate.IsEnabled = $true
    }
})

$Window.Add_Closing({ Disconnect-RoomProvisioningServices })

Show-Step -Step 1
[void]$Window.ShowDialog()
