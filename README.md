# Module: MailboxFolderSearch
A powershell module to assist creating targeted Compliance Content Search in Exchange Online mailboxes

## Theory

Exchange online does not support the old New-MailboxExportRequest and and New-MailboxImportRequest commandlets from Exchange on-premises anymore. The only official way to export and import mail data in Exchange Online is Outlook, which is extremely slow and more or less impossible to automate.
While we cannot work around the import with Outlook (except for moving the mailbox down to on-prem, if you still have an on-prem Exchange server, import and then move it back to the cloud), there is a way to speed up the export: Compliance Content Search

However, while ContentSearch is extremely fast for huge amounts of data, the more you want to narrow down the search the more complicated it gets.

## Targeted Search Queries

ContentSearch by default searches the complete mailbox. If you want to narrow it down to certain folders, you must enter the required folder ids as a KQL query.

**BUT**... you cannot just take the folder ids from Get-MailboxFolderStatistics. ContentSearch requires these Base64 encoded strings to be converted to a hex string!

**Example:**
The folder id of an inbox might look like this according to `Get-MailboxFolderStatistics`:
`FolderId : LgAAAABXV5oAt1XLR5mKDMdf/8tYAQA8g0kJjzXVTZlomtyxaMeJAAsEdkKpAAAB`
ContentSearch requires it like this:
`FolderQuery : folderid:3C8349098F35E54D99689ADCB168C789000B047642A90000`
_Note:_ the prefix folderid: is a keyword for KQL, not part of the hex-id.

While this might create some mixed feelings at the first glance, it has one big advantage: you can target folders in the mailbox's archive the same way! And to target more than one folder, simply concatenate the folder queries with a capitalized OR.

## The Module

To make things easier, use the custom module I linked to at the beginning of this section. It contains the following functions:

* `Convert-FolderIdToFolderQueryId`, to quickly convert a Base64 FolderId to the KQL hex format (without the prefix!)
* `Get-MFSFolderQueryId`, to get hex folder ids from a mailbox or archive including folderpath and foldertype
* `New-MFSComplianceSearchQuery`, to create the OR-joined folder query from a list of folderids
* `Start-MFSComplianceSearch`, to create and start a Compliance ContentSearch over a mailbox, it's archive, or both, maybe a selection of folders of your choice
* `Get-MFSComplianceSearchStatistics`, gets a Compliance Search by name or identity and expands the SearchStatistics from the returned JSON form to a more human readable array of objects
* `Expand-MFSComplianceSearchStatistics`, expands the SearchStatistics attribute returned from the regular `Get-ComplianceSearchfrom` its JSON format to a more human readable array of objects
* `Expand-MFSComplianceSearchActionResults`, expands the Results attribute returned from the regular `Get-ComplianceSearchAction` from its semicolon joined "Name: Value" list to a standard PSCustomObject

### Prerequisites

#### The module requires:

* Module ExchangeOnlineManagement to be loaded
* a working connection to ExchangeOnline
* a working connection to the Compliance Center
* Your code should look something like this:

```
# import the module
Import-Module ExchangeOnlineManagement
# connect to EXO
Connect-ExchangeOnline -UserPrincipalName username@agrana.net -ShowBanner:$false
# connect to Compliance Center
Connect-IPPSSession -UserPrincipalName username@agrana.net
```

_Note:_ make sure you use the one user for each connection, which provides the necessary right to do the job.


## HowTo

### The easy way, no details, please
```
Start-MFSComplianceSearch -Name 'MyCSearch' -PrimarySmtpAddress mc@fly.com -IncludeArchive -Interactive -InformationAction Continue
```
Will create and run a Compliance Content Search with the name "MyCSearch" over my mailbox AND archive (`-IncludeArchive`), give you an Out-GridView window, where you can choose the folders to search for (mark all folders you want and click ok) and add the chosen folders to the query. `-InformationAction Continue` will have the function post some informational messages during runtime.

Please check the functions help for details and additional parameters.

### I want more control
```
Get-EXOMailboxFolderStatistics mc@fly.com | Where-Object FolderType -eq 'User Created' | Convert-FolderIdToFolderQueryId | New-ComplianceSearchMailboxFolderQuery | Foreach-Object { New-ComplianceSearch -Name 'PSSearch mc@fly.com 20220411-112234' -ExchangeLocation 'mc@fly.com' -ContentMatchQuery $_}
```
* `Get-EXOMailboxFolderStatistics` (the ..EXO.. is the new graph based command, but works the same way as the old one without EXO) queries the folders from my mailbox (NOT the archive! This would requires a separate call using the `-Archive` parameter!)
* `Where-Objectfilters` for all folders of type "User Created", which are basically all mail folders excluding Inbox, plus a couple of contact folders
* `Convert-FolderIdToFolderQueryId` converts the folderids returned by `...FolderStatics` to the hex ids needed for the query
* `New-MFSComplianceSearchQuerycreates` the query string for KQL
* `New-ComplianceSearch` does not allow pipelining, so using a foreach-object we create the Compliance Search

### Exporting

Once the query has run, the resulting data can be exported to PST from the Compliance Center.
