<#
.SYNOPSIS
Convert a 64 char Base64 encoded Exchange FolderId to a 48 char hex string

.DESCRIPTION
Convert a 64 character base64 encoded Exchange Folder Id to a 48 character hex string Id, as it is needed for graph calls or ComplianceSearches (KQL)

.PARAMETER FolderId
base64 encoded FolderId, as returned from Get-(EXO)MailboxFolderStatistics

.EXAMPLE
Get-EXOMailboxFolderStatistics mc@fly.com | Convert-FolderIdToFolderQueryId

Returns the folder ids of all folders in mailbox mc@fly.com as hex ids.

.NOTES
2022-04-06 ... initial version by Maximlian Otter, base on the script in this article: https://docs.microsoft.com/en-us/microsoft-365/compliance/use-content-search-for-targeted-collections?view=o365-worldwide
2022-04-07 ... updated hex conversion code based on this article: https://gsexdev.blogspot.com/2019/01/converting-folder-and-itemids-in.html, which is ~10x faster
2022-04-11 ... made hex conversion function a standalone cmdlet, so it can be used without the Get-MailboxFolderQueryId function
#>
function Convert-FolderIdToFolderQueryId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string]
        $FolderId
    )

    process {
        if (![string]::IsNullOrEmpty($FolderId)) {
            $FolderIdConverted = [Convert]::FromBase64String($FolderId)
            $FolderIdConverted = [System.BitConverter]::ToString($FolderIdConverted[23..($FolderIdConverted.count-2)])
            $FolderIdConverted.Replace('-','')
        }
    }

}


<#
.SYNOPSIS
Get FolderIds ready to use with a compliance/content search query

.DESCRIPTION
Request the folder statistics from a give user and return the folders and folderids, the latter converted from the 64 char string to a 48 char format which can be used for Compliance and Content Searches.

The code requires the ExchangeOnlineManagement module to be loaded and a connection to ExchangeOnline must have been established.

.PARAMETER PrimarySmtpAddress
The PrimarySmtpAddress of the mailbox to query. But any alias of the mailbox should work.

.PARAMETER FolderName
Name or part of the name of a specific folder to query. More folder names can provided using commas ('one','two','three',...). A folder name will be returned if the given name appears anywhere within the actual folder name.

.PARAMETER Archive
A switch to return the folders from an archive mailbox instead of the primary

.EXAMPLE
Get-MailboxFolderQueryId -PrimarySmtpAddress marty@mcfly.com -FolderName 'Inbox','Doc' -Archive -Verbose -InformationAction Continue

Returns and converts the folder ids from the folders any folders containing "inbox" or "doc" in the archive mailbox of Marty McFly. Additional information and verbose output is created.

.NOTES
2022-04-06 ... initial version by Maximlian Otter, base on the script in this article: https://docs.microsoft.com/en-us/microsoft-365/compliance/use-content-search-for-targeted-collections?view=o365-worldwide
2022-04-07 ... updated hex conversion code based on this article: https://gsexdev.blogspot.com/2019/01/converting-folder-and-itemids-in.html, which is ~10x faster
#>
function Get-MFSFolderQueryId {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateScript({[bool][System.Net.Mail.MailAddress]::new($_)})]
        [string]
        $PrimarySmtpAddress,
        [Parameter()]
        [string[]]
        $FolderName,
        [Parameter()]
        [switch]
        $Archive
    )

    begin {

        Write-Information 'Checking Cmdlets'

        # use new EXO cmdlet if available
        if ([bool](Get-Command -Name Get-EXOMailboxFolderStatistics)) {
            $Get_MFS = 'Get-EXOMailboxFolderStatistics @args'
        # fall back to legacy cmdlet
        } elseif ([bool](Get-Command -Name Get-MailboxFolderStatistics)) {
            $Get_MFS = 'Get-MailboxFolderStatistics @args'
        # quit if none of the two is available
        } else {
            Throw 'Cmdlets "Get-EXOMailboxFolderStatistics" or "Get-MailboxFolderStatistics" not available.'
        }

        # enhance the scriptblock string with the -Archive parameter if requested
        if ($Archive) {
            $Get_MFS = $Get_MFS + ' -Archive'
        }
        $Get_MFS = [ScriptBlock]::Create($Get_MFS)

    }

    process {

        Write-Information "Querying folder statistics for `"$PrimarySmtpAddress`""
        $folderStatistics = & $Get_MFS -Identity $PrimarySmtpAddress

        if ($FolderName.Count -gt 0) {
            Write-Information "Filtering for requested folder name(s)"
            $folderStatistics = foreach ($fname in $FolderName) {
                Write-Verbose "Testing for folder name -like `"*$fname*`" ..."
                foreach ($folderStatistic in $folderStatistics) {
                    if ($folderStatistic.Name -like "*$fname*") {
                        $folderStatistic
                    }
                }
            }
        }

        Write-Information ("Calculating FolderSearchId for {0} folders" -f $folderStatistics.count)
        foreach ($folderStatistic in $folderStatistics) {
            $folderQueryId = Convert-FolderIdToFolderQueryId -FolderId $folderStatistic.FolderId
            [PSCustomObject]@{
                FolderPath = $folderStatistic.FolderPath
                FolderType = $folderStatistic.FolderType
                FolderQueryId = $folderQueryId
                FolderQuery = "folderid:$folderQueryId"
            }
        }

    }

}


<#
.SYNOPSIS
Creates a simple KQL query from a series of folder query ids.

.DESCRIPTION
The function takes a series of pre-converted folder ids (converted from base64 to hex format) checks them for validity, prefixes each id with "folderid:"
and OR-joins the individual folder queries to a single Keyword Query Lanugage query to be used in a ComplianceSearch.

.PARAMETER FolderQueryId
Mailbox folder id as 48 character hex string. Accepts multiple ids to be server as an array or via pipeline

.EXAMPLE
$Query = Get-EXOMailboxFolderStatistics mc@fly.com | Convert-FolderIdToFolderQueryId | New-MFSComplianceSearchQuery

.NOTES
2022-04-07 ... initial version by Maximilian Otter
#>
function New-MFSComplianceSearchQuery {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string[]]
        $FolderQueryId
    )

    begin {

        # internal function to make sure, the FolderQueryId looks legit
        function looksLikeFolderQueryId ([string]$HexString) {
            # must be 48 char long and represent a hex number
            if ($HexString.Length -eq 48) {
                $ishex = $true
                # 48 is a multiple of 16, so we don't have to calculate rest length and can hardcode a substring of length=16
                # the loop will therefor run max. 4 times, except if the check fails on an earlier substring
                for ($i = 0; $i -lt $HexString.Length -and $ishex; $i += 16) {
                    $substringishex = try { [System.Convert]::ToUInt64($HexString.Substring($i,16),16) -ge 0 } catch { $false }
                    $ishex = $ishex -and $substringishex
                }
                $ishex
            }
        }

        # prepare a generic list to collect the individual valid queries
        $FolderQuery = [System.Collections.Generic.List[string]]::new()

    }

    process {

        # since we accept arrays for the $FolderQueryId we have to treat it like multiple items and collect the result in an array
        # EXIT IF ONE FQID FAILS THE TEST! Because the result MUST contain all requested folders!
        $ValidQueryIds = foreach ($fqid in $FolderQueryId) {

            if (looksLikeFolderQueryId -HexString $fqid) {
                "folderid:$fqid"
            } else {
                Throw ("{0} does not look like a FolderQueryId (48 characters representing a hexadecimal number)." -f $fqid)
            }

        }

        # Add the resulting array to the generic list
        $FolderQuery.AddRange([System.Collections.Generic.List[string]]$ValidQueryIds)

    }

    end {

        # join the individual queries with a capitalized OR ( KQL requires capitalized operators! ) and return the resulting string
        $FolderQuery -join ' OR '

    }
}



<#
.SYNOPSIS
Create and run a ComplianceSearch over all or a manually chosen subset of folders in a mailbox

.DESCRIPTION
Create and run a ComplianceSearch over all or a manually chosen subset of folders in a mailbox. The search may target a mailbox' archive, too, or just the archive.

The code requires the ExchangeOnlineManagment module loaded and a connection to ExchangeOnline and the ComplianceCenter.

.PARAMETER Name
Name of the ComplianceSearch

If the name is not supplied, the ComplianceSearch will be named "PSSeach PrimarySMTPAddress yyyyMMdd-HHmmss", taking the PrimarySMTPAddress from the supplied parameter and the timestamp of the creation time.
The algorithm does not check, if the name is already taken. Using the timestamp down to the second should, however, be pretty safe.

.PARAMETER PrimarySmtpAddress
PrimarySMTPAddress of the mailbox to be searched. This can actually be any mail address the account can be identified with. The address will be used as the ExchangeLocation in the ComplianceSearch.
If there was no name supplied for the ComplianceSearch the PrimarySMTPAddress will be part of the auto-generated name.

.PARAMETER ArchiveOnly
A switch if you only want folders from the archive of the mailbox

.PARAMETER IncludeArchive
A switch if you want folders from the mailbox AND the archive

.PARAMETER Interactive
This will trigger an interactive Out-Gridview for the user to manually choose, which folder should be included in the search.

.EXAMPLE
Start-MFSComplianceSearch -Name "Archive Einstein" -PrimarySMTPAddress "albert@einstein.com" -ArchiveOnly

This will create and run a ComplianceSearch over all folders in the archive of mailbox albert@einstein.com.

.NOTES
2022-04-11 ... initial version by Maximilian Otter
#>
function Start-MFSComplianceSearch {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]
        $Name,
        [Parameter(Mandatory)]
        [Alias('Identity','ExchangeLocation')]
        [ValidateScript({[bool][System.Net.Mail.MailAddress]::new($_)})]
        [string]
        $PrimarySmtpAddress,
        [switch]
        $ArchiveOnly,
        [switch]
        $IncludeArchive,
        [switch]
        $Interactive
    )

    # Before anything else, check if the requested parameters make sense
    if ($ArchiveOnly -and $IncludeArchive) {
        Throw 'Switches "ArchiveOnly" and "IncludeArchive" cannot be used at the same time.'
    }

    # if no name was supplied for the ComplianceSearch, create one
    if (!$Name) {
        $Name = "PSSearch {0} {1}" -f $PrimarySmtpAddress,[datetime]::Now.ToString('yyyyMMdd-HHmmss')
    }

#region CHECK_ENVIRONMENT_AND_PREPARE_CMDLETS
    Write-Information 'Checking Cmdlets'

    # a connection to Compliance must have been established beforehand
    if (![bool](Get-Command New-ComplianceSearch -ErrorAction SilentlyContinue)) {
        Throw 'Cmdlet "New-ComplianceSearch" not available. Make sure the ExchangeOnlineManagement module is loaded and run "Connect-IPPSSession".'
    }

    # use new EXO cmdlet if available
    if ([bool](Get-Command -Name Get-EXOMailboxFolderStatistics -ErrorAction SilentlyContinue)) {
        $Get_MFS = 'Get-EXOMailboxFolderStatistics @args'
    # fall back to legacy cmdlet
    } elseif ([bool](Get-Command -Name Get-MailboxFolderStatistics -ErrorAction SilentlyContinue)) {
        $Get_MFS = 'Get-MailboxFolderStatistics @args'
    # quit if none of the two is available
    } else {
        Throw 'Cmdlets "Get-EXOMailboxFolderStatistics" or "Get-MailboxFolderStatistics" not available.'
    }
    $Get_MFStat = [ScriptBlock]::Create($Get_MFS)

    if ($ArchiveOnly -or $IncludeArchive) {
        # enhance the scriptblock string with the -Archive parameter if requested
        $Get_MFStatArc = [ScriptBlock]::Create(($Get_MFS + ' -Archive'))

    }
#endregion CHECK_ENVIRONMENT_AND_PREPARE_CMDLETS

#region GET_FOLDER_STATISTICS

    if (!$ArchiveOnly) {
        Write-Information 'Querying mailbox folders'
        $MbxFolders = & $Get_MFStat -Identity $PrimarySmtpAddress
    }

    if ($ArchiveOnly -or $IncludeArchive) {
        Write-Information 'Querying archive folders'
        $ArcFolders = & $Get_MFStatArc -Identity $PrimarySmtpAddress
    }

#endregion GET_FOLDER_STATISTICS

#region BUILD_KQL-Query

    if (!$Interactive) {
        Write-Information 'Creating search query'
        $Query = $MbxFolders + $ArcFolders | Where-Object FolderId | Convert-FolderIdToFolderQueryId | New-MFSComplianceSearchQuery

    # if Interactive was chosen, we change the retrieved objects so a user can distinguish between Mailbox and Archive folders
    # as well as only being presented the most necessary folder information
    } else {
        Write-Information 'Preparing data for manual interaction'
        $MbxFolders = foreach ($folder in $MbxFolders) {
            [PSCustomObject]@{
                Location = 'Mailbox'
                FolderType = $folder.FolderType
                Name = $folder.Name
                ItemsInFolder = $folder.ItemsInFolder
                FolderPath = $folder.FolderPath
                FolderId = $folder.FolderId
            }
        }
        $ArcFolders = foreach ($folder in $ArcFolders) {
            [PSCustomObject]@{
                Location = 'Archive'
                FolderType = $folder.FolderType
                ItemsInFolder = $folder.ItemsInFolder
                Name = $folder.Name
                FolderPath = $folder.FolderPath
                FolderId = $folder.FolderId
            }
        }
        Write-Information 'Creating search query'
        $Query = ($MbxFolders + $ArcFolders) | Out-GridView -PassThru | Convert-FolderIdToFolderQueryId | New-MFSComplianceSearchQuery

    }

#endregion BUILD_KQL-Query

#region CREATE-AND-RUN_COMPLIANCESEARCH

    Write-Information "Creating and running ComplianceSearch `"$Name`""
    New-ComplianceSearch -Name $Name -ExchangeLocation $PrimarySmtpAddress -ContentMatchQuery $Query | Start-ComplianceSearch
    Get-ComplianceSearch -Identity $Name

#region CREATE-AND-RUN_COMPLIANCESEARCH

}


<#
.SYNOPSIS
Get a Compliance Search by name or identity and expand the queries statitics

.DESCRIPTION
Get a Compliance Search by name and expand the queries statitics

.PARAMETER Identity
Name or Identity of the Compliance Search of which the statistics shouldbe shown

.EXAMPLE
Get-MFSComlianceSearchStatistics -Identity 'MySearch'

.NOTES
2022-04-12 ... initial version by Maximilian Otter
#>
function Get-MFSComplianceSearchStatistics {
    param (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Alias('Name')]
        [string]
        $Identity
    )

    process {
        Get-ComplianceSearch -Identity $Identity | Expand-MFSComplianceSearchStatistics
    }
}


<#
.SYNOPSIS
Expand the SearchStatistics JSON returned by Get-ComplianceSearch

.DESCRIPTION
Expand the SearchStatistics JSON returned by Get-ComplianceSearch

.PARAMETER SearchStatstics
Input the SearchStatistics attribute from Get-ComplianceSearch. It must be in JSON format.

.EXAMPLE
Get-ComplianceSearch -Identity 'MySearch' | Expand-MFSComplianceSearchStatistics

.NOTES
2022-04-12 ... initial version by Maximilian Otter
#>
function Expand-MFSComplianceSearchStatistics {
    param (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string]
        $SearchStatistics
    )

    process {

        $ExchangeBinding = ($SearchStatistics | ConvertFrom-JSON).ExchangeBinding
        ($ExchangeBinding.Search -as [array]) + $ExchangeBinding.Queries

    }
}


<#
.SYNOPSIS
Expand the ;-joined list of "Name: Value" pairs in the Get-ComplianceSearchAction Results attribute.

.DESCRIPTION
Expands the ;-joined list of "Name: Value" pairs from the Results attribute returned by Get-ComplianceSearchAction.
The result is returned as a standard PSCustomObject and the names are stripped off whitespace for easier followup processing

.PARAMETER Results
Takes the data from the Results attribute returned by Get-ComplianceSearchAction

.EXAMPLE
Get-ComplianceSearchAction -Identity 'MySearch' | Expand-MFSComplianceSearchActionResults

.NOTES
2022-04-12 ... initial version by Maximilian Otter
#>
function Expand-MFSComplianceSearchActionResults {
    param (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string]
        $Results
    )

    process {

        $ResultTable = $Results -split '; '

        $ResultHash = [ordered]@{}
        foreach ($row in $ResultTable) {
            $Parts = $row -split ': '
            $ResultHash.Add([cultureinfo]::CurrentCulture.TextInfo.ToTitleCase($Parts[0]).Replace(' ',''),$Parts[1])
        }

        [PSCustomObject]$ResultHash
    }
}
