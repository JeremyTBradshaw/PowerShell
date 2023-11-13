<#
    .SYNOPSIS
    Get all inbox rules that forward (including as an attachment) or redirect.

    .DESCRIPTION
    Outputs all Inbox Rules (a.k.a., Outlook rules) where ForwardTo, FowardAsAttachmentTo, or RedirectTrue are defined.
    Regardless of which action it is (or are, if multiple), an individual object is output for each target recipient.
    The output objects are commonized, and a property - ForwardOrRedirect - will show which type it is.  The 
    RoutingType property of the target recipient is the one which reveals whether it is an internal (RoutingType = 'EX')
    or external (RoutingType 'SMTP').
    
    The intent is to help with planning for external forwarding, as it relates to EOP Outbound spam protection policies,
    remote domains, and transport rules.  See the referenced Exchange Team blog post in the .LINK section

    .PARAMETER Identity
    Specifies the mailbox to check for forwarding/redirecting rules.

    .PARAMETER All
    Switch to indicate all mailboxes should be checked for forwarding/redirecting rules.

    .EXAMPLE
    Get-Mailbox -RecipientTypeDetails SharedMailbox | .\Get-ForwardingInboxRules.ps1 | fl

    .EXAMPLE
    .\Get-ForwardingInboxRules.ps1 -Identity Jeremy@jb365.ca

    .EXAMPLE
    $Rules = .\Get-ForwardingInboxRules.ps1 -All; $Rules | Export-Csv $Home\Desktop\FwdInboxRules.csv -NTI

    .LINK
    https://techcommunity.microsoft.com/t5/exchange-team-blog/all-you-need-to-know-about-automatic-email-forwarding-in/ba-p/2074888
#>
[CmdletBinding(DefaultParameterSetName = 'Identity')]
param (
    [Parameter(
        ParameterSetName = 'Identity',
        Position = 1,
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName
    )]
    [object[]]$Identity,
    [Parameter(ParameterSetName = 'All')]
    [switch]$All
)
begin {
    if ((Get-Command Get-Mailbox, Get-InboxRule -ErrorAction SilentlyContinue).Count -ne 2) {

        throw 'An active Exchange PowerShell session is required, along with access to the Get-Mailbox and Get-InboxRule cmdlets.'
    }
    $Script:startTime = [datetime]::Now
    $Script:stopwatchMain = [System.Diagnostics.Stopwatch]::StartNew()
    $Script:stopwatchPipeline = [System.Diagnostics.Stopwatch]::new()
    $Script:progress = @{
        Id              = 0
        Activity        = "$($PSCmdlet.MyInvocation.MyCommand.Name)"
        Status          = "Start time: $($startTime.ToString('yyyy-MM-ddTHH:mm:ss')) | Elapsed: $($stopWatchMain.Elapsed.ToString('hh\:mm\:ss'))"
        PercentComplete = -1
    }
    Write-Progress @progress

    if ($PSCmdlet.ParameterSetName -eq 'All') {
        try { $Script:Mailboxes = Get-Mailbox -ResultSize Unlimited -ea:Stop }
        catch { throw }
    }

    function getRules ([object[]]$mailbox) {
        try {
            foreach ($mbx in $mailbox) {
                $_rules = $null;
                $_rules = Get-InboxRule -Mailbox $mbx.PrimarySmtpAddress.ToString() -ea:SilentlyContinue |
                Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo }
        
                if ($_rules) { 
                    foreach ($_rule in $_rules) {
                        foreach ($_fwd in $_rule.ForwardTo) {
                            [PSCustomObject]@{
                                mbxDisplayName           = $mbx.DisplayName
                                mbxPrimarySmtpAddress    = $mbx.PrimarySmtpAddress
                                ruleIdentity             = $_rule.Identity
                                ruleEnabled              = $_rule.Enabled
                                ruleDescription          = $_rule.Description
                                ForwardOrRedirect        = 'Forward'
                                ruleForwardToAddr        = $_fwd.Address
                                ruleForwardToRoutingType = $_fwd.RoutingType
                            }
                        }
                        foreach ($_fwd in $_rule.ForwardAsAttachmentTo) {
                            [PSCustomObject]@{
                                mbxDisplayName           = $mbx.DisplayName
                                mbxPrimarySmtpAddress    = $mbx.PrimarySmtpAddress
                                ruleIdentity             = $_rule.Identity
                                ruleEnabled              = $_rule.Enabled
                                ruleDescription          = $_rule.Description
                                ForwardOrRedirect        = 'ForwardAsAttachment'
                                ruleForwardToAddr        = $_fwd.Address
                                ruleForwardToRoutingType = $_fwd.RoutingType
                            }
                        }
                        foreach ($_rdr in $_rule.RedirectTo) {
                            [PSCustomObject]@{
                                mbxDisplayName           = $mbx.DisplayName
                                mbxPrimarySmtpAddress    = $mbx.PrimarySmtpAddress
                                ruleIdentity             = $_rule.Identity
                                ruleEnabled              = $_rule.Enabled
                                ruleDescription          = $_rule.Description
                                ForwardOrRedirect        = 'Redirect'
                                ruleForwardToAddr        = $_rdr.Address
                                ruleForwardToRoutingType = $_rdr.RoutingType
                            }
                        }
                    }
                }
            }
        }
        catch { throw }
    }
    $stopWatchPipeline.Start()
    $Script:pipelineCounter = 0
}
process {
    $pipelineCounter++
    if ($PSCmdlet.ParameterSetName -eq 'Identity') { $Script:Mailboxes = if ($Identity[0].PrimarySmtpAddress) { $Identity[0] } else { try { Get-Mailbox $Identity[0] -ea:Stop } catch { throw } } }
    $Mailboxes | ForEach-Object {
        if ($stopWatchPipeline.ElapsedMilliseconds -ge 200) {

            $Script:progress.Status = "Start time: $($startTime.ToString('yyyy-MM-ddTHH:mm:ss')) | Elapsed: $($stopWatchMain.Elapsed.ToString('hh\:mm\:ss'))"
            $Script:progress.CurrentOperation = "Mailbox: $($_.DisplayName) ($($_.PrimarySmtpAddress))"
            Write-Progress @progress
            $_pct = if ($PSCmdlet.ParameterSetName -eq 'All') { (($pipelineCounter / $Mailboxes.Count) * 100) } else { -1 }
            Write-Progress -Activity 'Getting forward/redirect rules...' -Id 1 -ParentId 0 -PercentComplete $_pct
            $stopWatchPipeline.Restart()
        }
        try { getRules -mailbox $_ } catch { throw }
    }
}
end { Write-Progress @progress -Completed }  
