<#
    .SYNOPSIS
    Get all inbox rules that forward or redirect.

    .NOTES
    Quick draft for now.  To be refined later maybe.
#>
try {
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -ea:Stop
    foreach ($mbx in $Mailboxes) {

        $_rules = $null; $_rules = Get-InboxRule -Mailbox $mbx.PrimarySmtpAddress.ToString() -ea:SilentlyContinue |
        Where-Object { $_.ForwardTo -or $_.RedirectTo }
        if ($_rules) { 
            foreach ($_rule in $_rules) {
                [PSCustomObject]@{
                    mbxDisplayName        = $mbx.DisplayName
                    mbxPrimarySmtpAddress = $mbx.PrimarySmtpAddress
                    ruleIdentity          = $_rule.Identity
                    ruleEnabled           = $_rule.Enabled
                    ruleDescription       = $_rule.Description
                    ruleForwardTo         = $_rule.ForwardTo
                    ruleRedirectTo        = $_rule.RedirectTo
                }
            }
        }
    }
}
catch { throw }
