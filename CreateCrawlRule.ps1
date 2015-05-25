[CmdletBinding()]
Param(
  [Parameter(Mandatory=$False,Position=1)]
    [string]$environment="Default"
)

# Load Common Functions.
$commonScript = Join-Path $PSScriptRoot "\..\Common.ps1"
. $commonScript

StartLogging($MyInvocation.MyCommand.Definition)

try
{
    # Get xml file for this script.
    $directory = (Get-Item $MyInvocation.MyCommand.Definition).Directory.Name + "\"
    $xmlDoc = GetXmlFile (GetXmlFilename $MyInvocation.MyCommand.Definition) 'Enter xml filename' $directory

    # Search Service Application.
    $sa = $xmlDoc.SearchProperties.ServiceName
    $searchapp = GetServiceApplication $sa
    $crawlRules = GetXMLNodes $xmlDoc "SearchProperties" $environment "/CrawlRules/CrawlRule"    
    
    foreach ($crawlRule in $crawlRules)
    {
        $path = $crawlRule.Path

        $currentCrawlRule = Get-SPEnterpriseSearchCrawlRule -SearchApplication $searchApp | where { $_.Path -eq $path }
        if ($currentCrawlRule)
        {
            LogWarning "A Crawl Rule with path '$path' already exists. Delete the crawl rule and try again."
        }
        else
        {
            $type = $crawlRule.Type
            $authenticationType = $crawlRule.AuthenticationType            
            $accountName = $crawlRule.AccountName
            $accountPassword = Read-Host "Please enter the password for account: $($accountName)" -AsSecureString
        
            New-SPEnterpriseSearchCrawlRule -SearchApplication $searchapp -Path $path -Type $type -AccountName $accountName -AccountPassword $accountPassword -AuthenticationType $authenticationType

            Log "Created Crawl Rule with path '$path' for account '$accountName'"
        }
    }
}
catch
{    
    LogError $_.Exception.Message
}
finally
{
    StopLogging
}