# Welcome to my PowerShell repository

This repo is where I maintain much of my PowerShell scripting work.  If often seems to come back to Exchange for me and PowerShell, so much of my work here caters to that.  My only real boundary for what will live here though is that it's PowerShell (scripts, functions, modules, etc.), so over time the content will surely bounce around between more topics.

ℹ Most of my scripts / functions have detailed comment based help to make using them easier.  This can be read here on GitHub by viewing the code (which is syntax-highlighted for easy viewing), or directly in PowerShell which can also be handy.
```PowerShell
PS C:\> .\Get-MailboxLargeItems.ps1 -?
PS C:\> Get-Help Export-EXOMailbox.ps1 -Examples
PS C:\> Get-Help Get-MailboxTrusteeWebSQLEdition.ps1 -Full
```

# Project Portfolio

While the repo's content is an ever changing (mostly growing), list of items, this landing page/ReadMe is where I list some of my favorite ones, just some highlights per se.  Browse the code to see if there's anything else you might find useful.

## Currently In My Sights

I have been banging out reps with Microsoft Graph, particularly consuming it with [**MSGraphPSEssentials**](https://github.com/JeremyTBradshaw/MSGraphPSEssentials).  I have also been using SharePoint (Online) lists (via MS Graph/PowerShell) as an alternative to spreadsheets and databases.  It's quite fantastic, I must say.  With this, I have many ideas of solutions which would benefit from this same workflow.  One perfect candidate script which requires some updating is [**Get-EXOMailboxTrustee.ps1**](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-EXOMailboxTrustee.ps1); it needs to be updated to match the ouptput from [Get-MailboxTrustee.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrustee.ps1).  In addition to that, I think I will start writing some generic functions here to make using MS Graph and SharePoint lists specifically an easy and enjoyable process.  ...as time permits, but definitely on my radar.

## January 2022

### [Export-EXOMailbox.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Export-EXOMailbox.ps1)

This is a re-write of an old script by somebody from Microsoft who published it on TechNet Gallery as 'Export-Folder.ps1'.  It still does the essential same steps, however the code is cleaned up and made to support EXO/SCC PowerShell module of today.  Support has been added for selecting either or both the primary and archive mailbox, as well as for Inactive mailboxes.  Minus Security and Compliance Center PowerShell being a little flakey and failure-for-no-good-reason-prone, the script is quite slick.

### [Get-RecipientEmailAddresses.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-RecipientEmailAddresses.ps1)

Have you ever wanted to export all email addresses (a.k.a., proxyAddresses) so that each one was it's own object?  The use cases have stacked up for me over the years and finally here is a script to make the process nice and easy.  It includes some properties to make working with these in bulk easier (i.e., EmailAddress, Prefix, Domain, IsAcceptedDomain (optional)) and a couple properties to identify the recipient which holds the email address (PrimarySmtpAddress, Guid).  Some parameters have been added to help control which email addresses are included in the output, and if the addresses' domains should be compared against Exchange's list of Accepted Domains.  That last point helps to find email addresses that will prevent you from migrating your mailboxes to EXO (due to domain not being an Accepted Domain).

## 2021 Hiatus

Most of my focus in 2021 went to (work, and) [MSGraphPSEssentials](https://github.com/JeremyTBradshaw/MSGraphPSEssentials).  While you can bet PowerShell was open all day every day for the entire year, and that I was burning the midnight oil the entire time through, most of my work on stuff in this repo was refinements, updates, etc. to existing scripts.  You can see the commits and timestamps throughout the repo to see where my efforts went.

[Get-MailboxTrustee.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrustee.ps1) and [Get-MailboxTrusteeWebSQLEdition.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrusteeWebSQLEdition.ps1) got a nice chunk of TLC and the rest of the *MailboxTrustee* family of scripts was put to rest (deprecated).

## December 2020

### [Get-MailboxLargeItems.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxLargeItems.ps1) **V2**

I re-wrote this script and changed the way it goes about searching for items.  It borrows ideas from mainly Glen Scales' postings around the web.  I'm very happy with this script at this point and will publish it to the PowerShell Gallery soon.

### [EwsOAuthAppOnlyEssentials](https://github.com/JeremyTBradshaw/EwsOAuthAppOnlyEssentials)

**Update (2022-01-26)**: EwsOAuthAppOnlyEssentials has been deprecated and the same functionality ported over to [MSGraphPSEssentials](https://github.com/JeremyTBradshaw/MSGraphPSEssentials).

With my recent experience learning and using the EWS Managed API, I walked into something wonderful - I could reuse code from my MSGraphAppOnlyEssentials module to easily obtain EWS OAuth access tokens for use with Exchange Online!  Hence, the birth of this new module.  It's a near copy/paste of the functions `New-MSGraphAccessToken` (-> `New-EwsAccessToken`) and `New-SelfSignedMSGraphApplicationCertificate` (-> `New-SelfSignedEwsOAuthApplicationCertificate`).  Seems a little cheesy but I like it, and this module will also find it's way to the PowerShell Gallery soon.

## November 2020

### [Get-MailboxLargeItems.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxLargeItems.ps1) / [New-LargeItemsSearchFolder.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/New-LargeItemsSearchFolder.ps1)

I'm finally learning the EWS Managed API, starting at the last version - 2.2.  The first script, **Get-MailboxLargeItems.ps1** is my simplified rewrite of ['LargeItemChecks_2_2_0.ps1' from the TechNet Gallery](https://gallery.technet.microsoft.com/PowerShell-Script-Office-54d367ea).  That script has come up over the years for several EXO migration projects, and while it's great, I felt it was the perfect tool for me to learn EWS Managed API and that led to me wanting a simpler version of the same thing.  **New-LargeItemsSearchFolder.ps1** is the immediate next thing I had in mind once I was comfortable.  I had seen the idea over [here](http://www.flobee.net/search-mailboxes-for-large-items-that-may-impede-migrations-to-exchange-online/), and again, it was a great idea, but I found myself wanting my own rendition of something slightly different.
Review the comment-based help on these two scripts to see how to use them.

## September / October 2020

### [MSGraphAppOnlyEssentials 0.0.0.2](https://github.com/JeremyTBradshaw/MSGraphAppOnlyEssentials)

Ever since EXO v2 PS module came to life, and I retired **conex** (my Connect-Exchange (on-premises/EXO friendly) PS module), I had the itch to create a new PS module and have it published on the PowerShell Gallery.  I'm happy to say that itch has been scratched and I've created a module that makes using Microsoft Graph, in App-Only fashion using certificate credentials, a breeze.  Don't let the early version number scare you, both [**0.0.0.1**](https://www.powershellgallery.com/packages/MSGraphAppOnlyEssentials/0.0.0.1) and [**0.0.0.2**](https://www.powershellgallery.com/packages/MSGraphAppOnlyEssentials/0.0.0.2) have been put through their paces and are ready for action (though I'd start with 0.0.0.2 of course.

Check out the [README](https://github.com/JeremyTBradshaw/MSGraphAppOnlyEssentials/blob/master/README.md) to see how to use it, and get started.  MS Graph is like the jack of all trades, replacing many PowerShell modules that are required to meet the same functionality, and it goes way beyond that.

[**Available on the PowerShell Gallery**](https://www.powershellgallery.com/packages/MSGraphAppOnlyEssentials)

## March 2020

### [Get-EXOMailboxTrustee.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-EXOMailboxTrustee.ps1)

This is the EXO-exclusive successor to Get-MailboxTrustee.ps1.  It was written from scratch, rather than modified from the earlier script, and solely uses the new EXO V2 PS module cmdlets which consume the MS Graph / EXO REST API.  The hope is that this one will be able to handle org-wide reporting for even the biggest of EXO tenants.

It's worth noting, I won't be creating an EXO-exclusive successor to the Get-MailboxTrusteeWeb.ps1 lineup.  The use case for those was planning migration batches to EXO.  If/when offboarding from EXO becomes popular, I'll definitely tackle something like this then.

## November 2019

### [Search-InboxRuleChanges.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Search-InboxRuleChanges.ps1)

A recent addition, one that attempts to extend the built-in alerting capabilities of Microsoft 365's Security and Compliance Center (_SCC_) for potentially-malicious mailbox rule activity.  The problem being addressed, or at least attempting to be, is that only OWA-based activity is able to be further parsed for specifics to generate alerts for, in the M365 SCC.  **Search-InboxRuleChanges.ps1** further expands the dynamically nested multi-valued properties that are contained within the Unified Audit Log's AuditData proerty.  This in turn renders the _Outlook-based_ (Unified Audit Log activity: UpdateInboxRules) much more friendly to work with.  The only caveat is that you will need to schedule this check on your own (rather than set it up in SCC), and you'll also need to look after your own solution for getting alerts sent out (e.g. via email).

The `UseClientIPExcludedRanges` parameter ([bool] / $true by default) allows you to further scope the output to just changes made from non-recognized IP's.  Head on in and check it out.

## March / April 2019

### (R.I.P.) [ c o n e x ] : : To( "$(Exchange of your $Choice)" )
Conex has been retired since the Exchange Online PowerShell V2 module has been released and takes away half the need of conex in the first place.  It was fun, and might be back later for an on-premises Exchange exclusive.  Time will tell...

### [Get-MailboxTrustee.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrustee.ps1)

This script's target audience is anyone who needs to find mailbox permission relationships*, for example when planning a migration to Office 365.  The script outputs one object for every mailbox-trustee relationship, a format that works well for direct usability in a database or Excel table, from which pivot tables can very easily be crafted in seconds.

If you give it a try, try using it with -Verbose.  This, combined with its detailed progress, allows me to forego tracking with a _processesd.csv_ etc.

*Note: For Exchange Online, it'll be best to use Get-EXOMailboxTrustee.ps1 instead.

### [New-MailboxTrusteeReverseLookup.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/New-MailboxTrusteeReverseLookup.ps1)

New-MailboxTrusteeReverseLookup.ps1 is the sister script to Get-MailboxTrustee.ps1.  What to do after you've extracted your entire Exchange organization's list of mailbox & mailbox folder permissions (maybe even on a scheduled basis)?  Well, one thing you can do is respond to requests for mailbox permsission reverse lookups (e.g. __User__: _'What mailboxes do I have access to and what is my access level to each of them?_' __Answer__: '_Coming right up!_')

### [Get-MailboxTrusteeWebSQLEdition.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrusteeWebSQLEdition.ps1)

The SQL Edition is now at parity with the original version, and beyond (the database storage/engine really help tremendously).  If you have a large data set from Get-MailboxTrustee.ps1, this script has you covered with zippy speeds as it identifies webs.  The parameters allow for repeated trial and error / trial and cancel runs, as you fine tune the parameters to persuade the resulting web as desired..

While the original version (combined with the input-optimizing brother) works fine for small data sets, the SQL Edition is better suited to larger sets.  **Requirements?** SQL Express (only 2017 tested), and a database (default is TempDB) where current user has permissions to create/delete tables.

### [Get-MailboxUsage.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxUsage.ps1)

This script pulls some key details together to help determine how a mailbox is used.  Details include AD user account's last logon date, the mailbox's last logged on date / last logged on user, the newest item in the Inbox and Sent Items folders, and a few other items.  This one is arguably overlapping of Office/Microsoft 365's usage reports.  Its intended use is strictly with mailboxes, for tasks like determining if a UserMailbox is a good candidate for conversion to SharedMailbox, or Room/EquipmentMailbox.  Being that it can also be used with Exchange On-Premises _or_ EXO, it rather compliments the O/M365 usage reports.

## ... more to come / browse to see more.
