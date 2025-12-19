# Get-MailboxUsage.ps1

#Requires -Version 3
#Requires -Modules ActiveDirectory

[CmdletBinding()]

param(
    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [Alias(
        'Alias',
        'DistinguishedName',
        'Guid',
        'PrimarySmtpAddress',
        'SamAccountName')]
    [string]$Identity
)

begin {

    $StartTime = Get-Date
    Write-Debug -Message "begin {}.  Start time: $($StartTime.DateTime)"

    Write-Progress -Activity 'Get-MailboxUsage.ps1' -Status 'Initializing...' -SecondsRemaining -1 -CurrentOperation "Command: ""$($PSCmdlet.MyInvocation.Line)"""
    Write-Verbose -Message "Get-MailboxUsage.ps1 script begin.`n`tStart time: $($StartTime.DateTime)`n`tCommand: ""$($PSCmdlet.MyInvocation.Line)"""

    # Attempt to get a count of objects in the pipeline.

    $PipelineObjectCount = $null

    # Suppressing this one block's errors.
    try {
        $MyInvocationLineSplit = $PSCmdlet.MyInvocation.Line -split '\|'
        $MyInvocationLineSplit = $MyInvocationLineSplit -replace 'cls|Clear-Host',''

        $InvocationNameMatch = $MyInvocationLineSplit |
            Where-Object {$_ -like "*$($PSCmdlet.MyInvocation.InvocationName)*"}

        $InvocationNameIndex = $MyInvocationLineSplit.IndexOf($InvocationNameMatch)
        $MeasureableCommand = ($MyInvocationLineSplit | Select-Object -Index (0..($InvocationNameIndex - 1))) -join '|'
        $PipelineObjectCount = (Invoke-Expression -Command $MeasureableCommand | Measure-Object).Count
    }
    catch {<#Suppressed#>}

    if ($null -ne $PipelineObjectCount -and $PipelineObjectCount -gt 0) {

        $MailboxCounter = 0
        $MailboxPercentCompletePossible = $true
        Write-Verbose -Message "Successfully obtained pipeline object count ($($PipelineObjectCount)) for Write-Progress' -PercentComplete."
    }

    Write-Verbose -Message "Determining the connected Exchange environment."

    $ExPSSession = @()
    $ExPSSession += Get-PSSession |
    Where-Object {

        $_.ConfigurationName -eq 'Microsoft.Exchange' -and
        $_.State -eq 'Opened'
    }
    if ($ExPSSession.Count -eq 1) {
        $Exchange = $null

        switch ($ExPSSession.ComputerName) {

            outlook.office365.com {
                $Exchange = 'Exchange Online'
            }
            default {
                $Exchange = 'Exchange On-Premises'

                # Set scope to entire forest (important for multi-domain forests).
                Set-ADServerSettings -ViewEntireForest:$true
            }
        }
        Write-Verbose -Message "Connected environment is $($Exchange)."
    }
    else {
        Write-Warning -Message "Requires a *single* active (State: Opened) remote session to an Exchange server or EXO."
        break
    }


    # Define Write-Progress properties.
    $MainProgress = @{
        Activity = "Get-MailboxUsage.ps1 - Start time: $($StartTime.DateTime)"
        Status = "Working in $($Exchange) ($($ExPSSession.ComputerName))"
        Id = 0
        ParentId = -1
        SecondsRemaining = -1
    }
    Write-Progress @MainProgress
    Start-Sleep -Milliseconds 500

    $Progress1 = @{
        Activity = 'Get-Mailbox'
        Id = 1
        ParentId = 0
        SecondsRemaining = -1
    }
    Write-Progress @Progress1 -Status 'Ready'
    Start-Sleep -Milliseconds 500

    $Progress2 = @{
        Activity = 'Get-MailboxFolderStatistics'
        Id = 2
        ParentId = 0
        SecondsRemaining = -1
    }
    Write-Progress @Progress2 -Status 'Ready'
    Start-Sleep -Milliseconds 500

    $Progress3 = @{
        Activity = 'Get-MailboxStatistics'
        Id = 3
        ParentId = 0
        SecondsRemaining = -1
    }
    Write-Progress @Progress3 -Status 'Ready'
    Start-Sleep -Milliseconds 500

    $Progress4 = @{
        Activity = 'Get-ADUser'
        Id = 4
        ParentId = 0
        SecondsRemaining = -1
    }
    Write-Progress @Progress4 -Status 'Ready'
    Start-Sleep -Milliseconds 500

}

process {

try { # One try for the entire process block.  The idea is to skip over not-found mailboxes, but otherwise trudge on.


    # 1. Find the mailbox.


    Write-Progress @Progress1 -Status "Identity: $($Identity)"
    Write-Verbose -Message "[$(Get-Date -Format o)] Looking up mailbox with Identity ""$($Identity)""."

    $Mailbox = Get-Mailbox -Identity "$($Identity)" -ResultSize:1 -ErrorAction:Stop

    # Store the mailbox's DisplayName/PrimarySmtpAddress/Guid for the sake of ease through the rest of the script.
    [string]$mDisplay = $Mailbox.DisplayName
    [string]$mPSmtp   = $Mailbox.PrimarySmtpAddress
    [string]$mGuid    = $Mailbox.Guid

    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Mailbox lookup complete."

    if ($MailboxPercentCompletePossible) {

        Write-Verbose -Message "[Mailbox: $($mPSmtp)] Index # $($MailboxCounter) of $($PipelineObjectCount - 1) ($($PipelineObjectCount) total)."

        $MailboxCounter++
        Write-Progress @MainProgress -CurrentOperation "Processing mailbox #$MailboxCounter of $PipelineObjectCount`: $($mDisplay) ($($mPSmtp))" -PercentComplete (($MailboxCounter/$PipelineObjectCount) * 100)
    }
    else {
        Write-Progress @MainProgress -CurrentOperation "Processing mailbox: $($mDisplay) ($($mPSmtp))"
    }


    # 2. Get-MailboxFolderStatistics (newest Inbox/Sent Items items).


    Write-Progress @Progress2 -Status 'Working'
    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Getting Inbox and Sent Items folder statistics (Get-MailboxFolderStatistics)."

    $InboxStats = Get-MailboxFolderStatistics -Identity "$($mGuid)" -FolderScope:Inbox -IncludeOldestAndNewestItems | Select-Object -First 1
    $NewestInboxItem = $null
    $NewestInboxItem = try { $InboxStats[0].NewestItemLastModifiedDate.ToShortDateString() } catch {}

    $SentItemsStats = Get-MailboxFolderStatistics -Identity "$($mGuid)" -FolderScope:SentItems -IncludeOldestAndNewestItems | Select-Object -First 1
    $NewestSentItemsItem = $null
    $NewestSentItemsItem = try { $SentItemsStats[0].NewestItemLastModifiedDate.ToShortDateString() } catch {}

    Write-Progress @Progress2 -Status 'Ready'
    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Finished getting Inbox and Sent Items folder statistics."


    # 3. Get-MailboxStatistics (LastLogonTime, LastLoggedOnUser).


    Write-Progress @Progress3 -Status 'Working'
    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Getting mailbox's last logged on date/user (Get-MailboxStatistics)."

    $MailboxStats = Get-MailboxStatistics -Identity "$($mGuid)" -ErrorAction:SilentlyContinue
    $MailboxLastLoggedOnDate = $null
    $MailboxLastLoggedOnDate = try {$MailboxStats.LastLogonTime.ToShortDateString() } catch {}

    Write-Progress @Progress3 -Status 'Ready'
    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Finished getting mailbox's last logged on date/user."


    # 4. Get-ADUser (Enabled?, LastLogonDate).


    Write-Progress @Progress4 -Status 'Working'
    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Getting AD user's enabled status and last logon date (Get-ADUser)."

    $SearchableExchangeGuid = '\' +
                            (($Mailbox.ExchangeGuid.Guid -replace '-' -split '(..)' |
                              Where-Object {$_})[3,2,1,0,5,4,7,6,8,9,10,11,12,13,14,15] -join '\')

    $ADUser = Get-ADUser -Filter "msExchMailboxGuid -eq '$($SearchableExchangeGuid)'" -Properties LastLogonDate -ErrorAction:SilentlyContinue
    $ADUserLastLogonDate = $null
    $ADUserLastLogonDate = try { $ADUser.LastLogonDate.ToShortDateString() } catch {}

    Write-Progress @Progress4 -Status 'Ready'
    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Finished getting AD user's enabled status and last logon date."


    # 5. Combine and output properties as a PSCustomObject.


    Write-Debug -Message "[Mailbox: $($mPSmtp)] We are ready to output our combined properties object.  Inspectable variables: `$Mailbox, `$InboxStats, `$SentItemsStats, `$MailboxStats, `$ADFilter, `$ADUser"
    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Creating and outputting combined properties object."

    $CombinedObject = [PSCustomObject]@{

                    DisplayName = $mDisplay
             PrimarySmtpAddress = $mPSmtp
                           Guid = $mGuid
        MailboxLastLoggedOnDate = $MailboxLastLoggedOnDate
        MailboxLastLoggedOnUser = $Mailboxstats.LastLoggedOnUserAccount
                NewestInboxItem = $NewestInboxItem
            NewestSentItemsItem = $NewestSentItemsItem
                        Enabled = $ADUser.Enabled
                  LastLogonDate = $ADUserLastLogonDate
    }
    Write-Output -InputObject:$CombinedObject

    Write-Verbose "[$(Get-Date -Format o)] Finished with Identity ""$($Identity)""."

} # end try

catch {

    Write-Warning -Message "Failed on Identity ""$($Identity)""."
    Write-Warning -Message "Error:`n$($error[0])"
    Write-Warning -Message 'Moving onto next item (if applicable).'

}

} # end process

end {

    $EndTime = Get-Date
    Write-Debug -Message "We are here -->: end {}.  End time: $($EndTime.DateTime)"
    Write-Verbose -Message "Get-MailboxUsage.ps1 script end.`n`tEnd time: $($EndTime.DateTime)"

    Write-Progress @Progress4 -Completed
    Write-Progress @Progress3 -Completed
    Write-Progress @Progress2 -Completed
    Write-Progress @Progress1 -Completed
    Write-Progress @MainProgress -Completed

}

<#

    .Synopsis

    Check mailbox for:

    - Newest item in Inbox and Sent Items folders.
    - Mailbox's LastLoggedOnDate / LastLoggedOnUser.
    - Mailbox's user account's LastLogonDate.

    .Parameter Identity [string]

    Accepts pipeline input directly or by property name (caution not to pipe
    directly from Exchange cmdlets to avoid concurrent/busy pipeline errors
    (piping single objects is OK).  Instead, first store multiple objects in a
    variable, then pipe the variable (see examples).

    Accepted properties from the pipeline are (because other properties have
    proven to be failure-prone):

    - SamAccountName
    - Alias
    - PrimarySmtpAddress
    - Guid
    - DistinguishedName

    .Link

    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxUsage.ps1

#>
