[CmdletBinding()]
Param(
  [Parameter(Mandatory=$False,Position=1)]
    [string]$environment="Default"
)

# Load Common Functions.
$Global:scriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$commonScript = Join-Path $Global:scriptPath "Common.ps1"
. $commonScript

StartLogging($MyInvocation.MyCommand.Definition)

try
{
    # Get xml file for this script.
    $xmlDoc = GetXmlFile(GetXmlFilename $MyInvocation.MyCommand.Definition)
    
    # Variables.
    $siteUrl = GetSiteUrl($environment)
    $searchAppName = $xmlDoc.ContentSources.SearchServiceApplicationName
    $searchServiceApp = GetServiceApplication $searchAppName
    $proxyGroup = Get-SPServiceApplicationProxyGroup -default
    $contentSources = GetXMLNodes $xmlDoc "ContentSources" $environment "/ContentSource"

    foreach ($contentSource in $contentSources)
    {	
        $contentSourceName = $contentSource.Name
        $lobSystemSet = @($contentSource.LOBSystemName, $contentSource.LOBSystemInstanceName)

	    Log "Attempting to create new content source '$contentSourceName'"	    
	
	    #################################
	    ### Create the content source ###
	    #################################
	    $existingCS = Get-SPEnterpriseSearchCrawlContentSource -Identity $contentSourceName -SearchApplication $searchServiceApp -ErrorAction:SilentlyContinue
	    if(-not ($existingCS -eq $null)) { throw "Content source '$contentSourceName' already exists!" }
	
        New-SPEnterpriseSearchCrawlContentSource -Name $contentSourceName -SearchApplication $searchServiceApp -Type $contentSource.Type -LOBSystemSet $lobSystemSet -BDCApplicationProxyGroup $proxyGroup

        $createdCS = Get-SPEnterpriseSearchCrawlContentSource -Identity $contentSourceName -SearchApplication $searchServiceApp -ErrorAction:SilentlyContinue
	    if(-not ($createdCS -eq $null)) { 
		    Log "Successfully created content source '$contentSourceName'" 
            
            $createdCS | Set-SPEnterpriseSearchCrawlContentSource -ScheduleType Full -DailyCrawlSchedule -CrawlScheduleRunEveryInterval 1
		
		    if($contentSource.RunFullCrawl -eq $true) {
			    Log "Starting Full Crawl of content source '$contentSourceName'"
			    $createdCS.StartFullCrawl()
			    Log "Full Crawl started."
		    }
	    }
	    else { 
		    throw "Failed to create content source!" 
	    }
    }
}
catch [Exception] {
    LogError $_.Exception.Message
}
finally
{
    StopLogging
}