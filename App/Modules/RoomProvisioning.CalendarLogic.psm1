#requires -Version 5.1

function Get-StandardCalendarProcessingParams {
    <#
        The "Standard mode" defaults - identical to what the original
        script always applied, kept as one place so Custom mode can start
        from the same baseline and only override what the user answered.
    #>
    [CmdletBinding()]
    param()

    @{
        AutomateProcessing             = 'AutoAccept'
        AllowRecurringMeetings         = $true
        DeleteAttachments              = $true
        ProcessExternalMeetingMessages = $true
        RemovePrivateProperty          = $false
        AddAdditionalResponse          = $true
        AdditionalResponse             = 'This is a Microsoft Teams Meeting room!'
        AddOrganizerToSubject          = $false
        DeleteSubject                  = $false
        DeleteComments                 = $false
        AllowConflicts                 = $false
    }
}

function ConvertTo-CalendarProcessingParams {
    <#
        Translates the plain-language answers collected in Custom mode into
        the Set-CalendarProcessing parameter set.

        Cleanup vs. the original script: AddOrganizerToSubject is always
        kept $false (a Teams Rooms panel has no real use for the organizer
        name being stuffed into the subject line, private or not) - so the
        "private room" answer only ever touches DeleteSubject/DeleteComments,
        instead of the original script's inconsistent mix of the three.

        $Answers is expected to have these properties (all optional except
        IsPrivate/AllowConflictingSeries/BookingWindowDays/RequireApproval -
        anything else left unset keeps the standard-mode default):
          IsPrivate                 [bool]
          AllowConflictingSeries    [bool]
          ConflictPercentageAllowed [int]    (only used if AllowConflictingSeries)
          MaximumConflictInstances  [int]    (only used if AllowConflictingSeries)
          BookingWindowDays         [int]
          RequireApproval           [bool]
          ApprovalDelegates         [string[]] (only used if RequireApproval)
          AllowRecurringMeetings    [bool]
          RemoveAttachments         [bool]
          AllowExternalRequests     [bool]
          RemovePrivateFlag         [bool]
          AdditionalResponseText    [string] (blank/omitted = no auto-reply text added)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Answers
    )

    $params = Get-StandardCalendarProcessingParams

    if ($Answers.PSObject.Properties['IsPrivate'] -and $Answers.IsPrivate) {
        $params.DeleteSubject  = $true
        $params.DeleteComments = $true
    }

    if ($Answers.PSObject.Properties['AllowConflictingSeries']) {
        $params.AllowConflicts = [bool]$Answers.AllowConflictingSeries
        if ($Answers.AllowConflictingSeries) {
            $params.ConflictPercentageAllowed = [int]$Answers.ConflictPercentageAllowed
            $params.MaximumConflictInstances   = [int]$Answers.MaximumConflictInstances
        }
    }

    if ($Answers.PSObject.Properties['BookingWindowDays'] -and $Answers.BookingWindowDays) {
        $params.BookingWindowInDays = [int]$Answers.BookingWindowDays
    }

    if ($Answers.PSObject.Properties['RequireApproval'] -and $Answers.RequireApproval) {
        # Requests from people in ResourceDelegates' "trusted" list still
        # auto-accept; everyone else needs a delegate to approve. Setting
        # AllBookInPolicy/AllRequestOutOfPolicy to $false is what actually
        # turns approval on - AutomateProcessing stays AutoAccept.
        $params.AllBookInPolicy      = $false
        $params.AllRequestOutOfPolicy = $false
        if ($Answers.PSObject.Properties['ApprovalDelegates'] -and $Answers.ApprovalDelegates) {
            $params.ResourceDelegates = @($Answers.ApprovalDelegates)
        }
    }

    if ($Answers.PSObject.Properties['AllowRecurringMeetings']) {
        $params.AllowRecurringMeetings = [bool]$Answers.AllowRecurringMeetings
    }

    if ($Answers.PSObject.Properties['RemoveAttachments']) {
        $params.DeleteAttachments = [bool]$Answers.RemoveAttachments
    }

    if ($Answers.PSObject.Properties['AllowExternalRequests']) {
        $params.ProcessExternalMeetingMessages = [bool]$Answers.AllowExternalRequests
    }

    if ($Answers.PSObject.Properties['RemovePrivateFlag']) {
        $params.RemovePrivateProperty = [bool]$Answers.RemovePrivateFlag
    }

    if ($Answers.PSObject.Properties['AdditionalResponseText']) {
        if ([string]::IsNullOrWhiteSpace($Answers.AdditionalResponseText)) {
            $params.AddAdditionalResponse = $false
            $params.Remove('AdditionalResponse')
        } else {
            $params.AddAdditionalResponse = $true
            $params.AdditionalResponse    = $Answers.AdditionalResponseText
        }
    }

    return $params
}

function Set-RoomCalendarProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][hashtable]$CalendarParams
    )

    Set-CalendarProcessing -Identity $Identity @CalendarParams -ErrorAction Stop
}

Export-ModuleMember -Function Get-StandardCalendarProcessingParams, ConvertTo-CalendarProcessingParams, Set-RoomCalendarProcessing
