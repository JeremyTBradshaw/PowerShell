<#
    .SYNOPSIS
    Export an Exchange Online mailbox to <TBD>...

    .DESCRIPTION
    The goal of this script is to use the official Microsoft Graph PowerShell SDK modules to export mailbox contents,
    including folder structure, to OneDrive.  The exact details are TBD as I eventually work my way through it.
    I expect many options will be possible to offer over time, such as:

      - Exporting to OneDrive, SharePoint, local file system, etc.
      - Individual emails/items, specific folders, or entire mailbox (primary/archive).
      - Assuming EML files.
      - Maybe headers only option.
      - Open to suggestions (please use GitHub Discussions).
#>
#Requires -Version 5.1
#Requires -Module @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion ='2.0.0'; Guid ='883916f2-9184-46ee-b1f8-b6a2fb784cee' }
#Requires -Module @{ ModuleName = 'Microsoft.Graph.Files'; ModuleVersion ='2.0.0'; Guid ='45ddab16-496a-4ef0-ac17-dbf0f93494d3' }
#Requires -Module @{ ModuleName = 'Microsoft.Graph.Mail'; ModuleVersion ='2.0.0'; Guid ='6e4d36b5-7ff2-454b-8572-674b3ab0362b' }
#Requires -Module @{ ModuleName = 'Microsoft.Graph.Sites'; ModuleVersion ='2.0.0'; Guid ='7ae8c25b-f1dd-466d-a022-b5489f919c70' }
[CmdletBinding()]
param ( )
begin {}
process {}
end {}
