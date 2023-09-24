<#
    .SYNOPSIS
    Get all Stream Classic Webparts from all sites in tenant.

    .DESCRIPTION
    This script gets all Stream Classic Webparts from all sites in tenant and exports the result to a CSV file.
    The CSV file can be exported to an Azure Storage Account, if needed.

    .PARAMETER SiteGuidForSingleSite
    If this parameter is provided, the script will query this site only. Otherwise, all sites in tenant will be queried.

    .PARAMETER PageSize
    The number of sites to query in one request. Default: 200

    .PARAMETER ResultSize
    The maximum number of sites to query. Default: $null (all sites)

    .PARAMETER ExportToAzureStorage
    If this switch is provided, the CSV file will be exported to an Azure Storage Account. Default: $false

    .PARAMETER ResourceGroupName
    The name of the resource group, where the Azure Storage Account is located. Default: "StreamClassicWebpart"

    .PARAMETER StorageAccountName
    The name of the Azure Storage Account, where the CSV file will be stored. Default: "sacsvfilestreamclassic"

    .PARAMETER StorageContainerName
    The name of the Azure Storage Container, where the CSV file will be stored. Default: "stream-classic-webparts"

    .PARAMETER MaxRetryCount
    The maximum number of retries, if the Graph API returns an error. Default: 3

    .PARAMETER CustomVerbose
    If this switch is provided, the script will output more information. Default: $false

    .OUTPUTS
    The script outputs a CSV file with the following columns:
    - SiteName
    - SiteUrl
    - SiteId
    - SiteOwner
    - PageName
    - PageId
    - WebpartTitle
    - EmbedCode

    .NOTES
    Prerequisites
        - Azure Resources
            - Azure Function App (Premium Plan)
                - Managed Identity
				    - Add "Microsoft.Graph\Sites.Read.All" permission
                    - Add "Storage Account Contributor" permission for storage account, where the Csv file will be stored
                - Upload this PowerShell Code in new timer triggered function
            - required Modules (Microsoft.Graph.Beta.Sites not possible, as Get-MgBetaSitePageWebpart is not available in Graph module)
                - Microsoft.Graph.Authentication (2.*)
                - Az.Accounts (2.*)
                - Az.Storage (5.*)

    .EXAMPLE
    Get-StreamClassicWebpart.ps1
#>

# Input bindings are passed in via param block.
param($Timer)

# region Variables

# If this parameter is provided, the script will query this site only. Otherwise, all sites in tenant will be queried.
$SiteGuidForSingleSite = ''

# The number of sites to query in one request
$PageSize = 200

# The maximum number of sites to query
$ResultSize = 5000

# If this switch is provided, the CSV file will be exported to an Azure Storage Account
$ExportToAzureStorage = 0 # default: $false

# The name of the resource group, where the Azure Storage Account is located
$ResourceGroupName = "StreamClassicWebpart"

# The name of the Azure Storage Account, where the CSV file will be stored
$StorageAccountName = "sacsvfilestreamclassic"

# The name of the Azure Storage Container, where the CSV file will be stored
$StorageContainerName = "stream-classic-webparts"

# The maximum number of retries, if the Graph API returns an error
$MaxRetryCount = 3

# If this switch is provided, the script will output more information
$CustomVerbose = 0 # default: $false, ordinary Verbose parameter is not visible in Azure Automation

# endregion Variables

# region Functions

function Invoke-CommandRetry {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Url,

        [Int]
        $MaxRetryCount
    )

    # Variables
    $throttlingRetry = 0
    $tokenRetry = 0
    $leaveInnerLoop = $false
    
    # Loop until the query was successful or the maximum number of retries is reached
    do {
        try {
            $result = Invoke-MgGraphRequest -Method 'GET' -Uri $Url
        }
        catch {
            # end script, if maximum number of retries is reached
            if (($tokenRetry -eq $MaxRetryCount) -or ($throttlingRetry -eq $MaxRetryCount)) {
                if ($tokenRetry -eq $MaxRetryCount) {
                    Write-Error "Unauthorized: Tried to reconnect $($tokenRetry) times without success. Exiting."
                }
                if ($throttlingRetry -eq $MaxRetryCount) {
                    Write-Error "Custom throttling: Tried to wait $($throttlingRetry) times without success. Exiting."
                }
                $PSCmdlet.ThrowTerminatingError($_)
            }
            # Throttling: if the error message contains "retry again later", wait and retry
            if ($_ -like "*retry again later*") {
                $throttlingRetry++
                $waitTime = 5 * [math]::Pow($throttlingRetry, $throttlingRetry)
                Write-Error "Custom throttling: Waiting for $waitTime seconds before retrying. $throttlingRetry. attempt"
                $null = Start-Sleep -Seconds $waitTime
                continue
            }
            # Unauthorized: if the error message contains "Authentication needed", try to get a new token
            elseif ($_ -like "*Authentication needed*") {
                $tokenRetry++
                Write-Error "Unauthorized: Trying to get new token. $tokenRetry. attempt"
                $null = Connect-MgGraph -Identity -NoWelcome -ErrorAction SilentlyContinue
                continue
            }
            # Conversion error - skipping this site
            elseif ($_ -like "*Cannot convert the JSON string because a dictionary that was converted from the string contains the duplicated key 'title'*") {
                Write-Error "Cannot convert the JSON string because a dictionary that was converted from the string contains the duplicated key 'title'. Skipping..."
                $errorAt = [PSCustomObject]@{
                    "SiteName"     = $site.name
                    "SiteUrl"      = $site.webUrl
                    "SiteId"       = $site.id
                    "SiteOwner"    = ($siteOwner.owner | ConvertTo-Json -Compress)
                    "PageName"     = $page.name
                    "PageId"       = $page.id
                    "WebpartTitle" = $sitePageWebpart.data.title
                    "EmbedCode"    = $sitePageWebpart.data.properties.embedCode
                    }
                Write-Error $errorAt
            }
            # other error: end script
            else {
                Write-Error $_
                $PSCmdlet.ThrowTerminatingError($_)
            }
        }

        # if the query was successful, leave the inner loop
        $leaveInnerLoop = $true

    } until($leaveInnerLoop)
    
    # return result
    $result
}

# endregion Functions

# region Execution

# Connect to Graph API (needed for Graph queries)
Connect-MgGraph -Identity -NoWelcome

# Get all sites (excluding OneDrive sites)
$sitesUrl = @'
/beta/sites/getAllSites?$select=id,name,webUrl&$filter=not(contains(webUrl,'-my.sharepoint.com'))&$orderBy=name asc&$top={0}
'@ -f $PageSize

# if SiteGuid is provided, query this site only
if ($SiteGuidForSingleSite) {
    $sitesUrl = '/beta/sites/{0}?$select=id,name,webUrl' -f $SiteGuidForSingleSite
}

$outputDate = Get-Date -Format yyyyMMddhhmmss
$loopCounter = 0
$fileCounter = 0
do {
    # New start of this loop
    Write-Host "--- Working on next set of sites ($($loopCounter * $PageSize) - $(($loopCounter + 1) * $PageSize)) ---"
    $localOutputPath = Join-Path -Path $env:TEMP -ChildPath ("$outputDate-GetStreamClassicWebpart-{0}.csv" -f $fileCounter)
    $fileCreated = $false

    # Query sites
    if ($CustomVerbose) { Write-Host "VERBOSE: Querying $sitesUrl" }
    $sites = Invoke-CommandRetry -Url $sitesUrl -MaxRetryCount $MaxRetryCount
    if (($sites.value.Count -eq 0) -and ($null -ne $sites)) {
        # if there is only one site, the value property is not an array...
        $sitesValue = $sites
        if ($CustomVerbose) { Write-Host "VERBOSE: Found 1 site." }
    }
    else {
        # ... otherwise use the value property in the following loop
        $sitesValue = $sites.value
        if ($CustomVerbose) { Write-Host "VERBOSE: Found $($sitesValue.Count) sites." }
    }

    if($sitesValue.Count -eq 0) {
        continue
    }

    # Loop through sites
    foreach ($site in $sitesValue) {
        
        # Skip sites without name at the REST query (for example 'search' site has no value in name property)
        if ($null -eq $site.name) {
            continue
        }

        # Query pages of site
        $sitePagesUrl = '/beta/sites/{0}/pages/microsoft.graph.sitePage?$select=id,name' -f $site.id
        if ($CustomVerbose) { Write-Host "VERBOSE: Querying $sitePagesUrl" }
        $sitePages = Invoke-CommandRetry -Url $sitePagesUrl -MaxRetryCount $MaxRetryCount
        if ($CustomVerbose) { Write-Host "VERBOSE: Found $($sitePages.value.Count) pages." }

        if($sitePages.value.Count -eq 0) {
            continue
        }

        # Loop through pages
        foreach ($page in $sitePages.value) {

            # Query webparts of page
            $sitePageWebpartsUrl = '/beta/sites/{0}/pages/{1}/microsoft.graph.sitePage/webparts' -f $site.id, $page.id
            if ($CustomVerbose) { Write-Host "VERBOSE: Querying $sitePageWebpartsUrl" }
            $sitePageWebparts = Invoke-CommandRetry -Url $sitePageWebpartsUrl -MaxRetryCount $MaxRetryCount
            if ($CustomVerbose) { Write-Host "VERBOSE: Found $($sitePageWebparts.value.Count) webparts." }

            if($sitePageWebparts.value.Count -eq 0) {
                continue
            }
            
            # Loop through webparts
            foreach ($sitePageWebpart in $sitePageWebparts.value) {

                if ($CustomVerbose) { Write-Host "VERBOSE: Checking web part $($sitePageWebpart.data.title)." }

                if ($sitePageWebpart.data.properties.embedCode -notlike "*https://web.microsoftstream.com/*") {
                    if ($CustomVerbose) { Write-Host "VERBOSE: No Stream Classic web part found" }
                    continue
                }

                # Get site owner
                $siteOwnerUrl = '/beta/sites/{0}/drive?$select=owner' -f $site.id
                if ($CustomVerbose) { Write-Host "VERBOSE: Querying $siteOwnerUrl" }
                $siteOwner = Invoke-CommandRetry -Url $siteOwnerUrl -MaxRetryCount $MaxRetryCount

                if($siteOwner.owner.count -eq 0) {
                    continue
                }

                # Stream Classic webpart found
                $currentObject = [PSCustomObject]@{
                    SiteName     = $site.name
                    SiteUrl      = $site.webUrl
                    SiteId       = $site.id
                    SiteOwner    = ($siteOwner.owner | ConvertTo-Json -Compress)
                    PageName     = $page.name
                    PageId       = $page.id
                    WebpartTitle = $sitePageWebpart.data.title
                    EmbedCode    = $sitePageWebpart.data.properties.embedCode
                }

                if ($CustomVerbose) { Write-Host "VERBOSE: STREAM CLASSIC FOUND!" }
                Write-Host $currentObject

                $null = $currentObject | Export-Csv -Path $localOutputPath -Encoding utf8 -Delimiter ';' -Append
                $fileCreated = $true
            }
        }
    }

    if($ExportToAzureStorage -and $fileCreated) {
        # Copy local output file to Azure Storage
        # "StorageAccount Contributor" role is needed for the Automation Account Managed Identity!
        $storageContext = (Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName).Context
        $null = Set-AzStorageBlobContent -Container $StorageContainerName -File $localOutputPath -Context $storageContext -Force
        $fileCounter++
    }

    # Get next page of sites, if available. Otherwise, set $sitesUrl to $null automatically to end the loop
    $sitesUrl = $sites.'@odata.nextlink'
    $loopCounter++

    # Check if the maximum number of sites is reached
    if($ResultSize -and ($loopCounter * $PageSize -ge $ResultSize)) {
        $sitesUrl = $null
    }
}
while ($null -ne $sitesUrl)

# endregion Execution