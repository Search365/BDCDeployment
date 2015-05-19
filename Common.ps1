set-ExecutionPolicy RemoteSigned
Add-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue

<###################################################################
#  Gets the script path from the command path.
###################################################################>
function Global:GetScriptPath()
{
    return $Global:scriptPath
}

<###################################################################
#  Logs the message and provides datetime/user info.
###################################################################>
function Global:Log ($message, $foregroundColor="green") {
    [string] $date = get-date -uformat "%G-%m-%d %T"
	Write-Host "[$date] " -ForegroundColor $foregroundColor -nonewline
	[string] $thisHost = ($env:COMPUTERNAME+'.'+$env:USERDNSDOMAIN).toLower()
	Write-Host "[$thisHost] " -ForegroundColor $foregroundColor -nonewline
	Write-Host "".padright(4) -Foregroundcolor $foregroundColor -nonewline
	Write-Host -ForegroundColor $foregroundColor "$message`r`n"
}

function Global:LogWarning($message) {
    Log $message "yellow"
}

function Global:LogError($message) {
    Log $message "red"
}

<###################################################################
#  Terminates the current script
###################################################################>
function Global:TerminateScript {
	param($message)	
	LogError($message)
	exit(1)
}

<###################################################################
#  Gets the xml file specified from the current directory. If not 
#  found it prompts the user to enter the filename. Appends the 
#  directoryPath to the current directory if specified.
###################################################################>
function Global:GetXmlFile($filename, $userMessage = 'Enter xml filename', $directoryPath='')
{
    $filename = GetFilePath $filename $directoryPath
    $fileExists = Test-Path $filename

    if (!$fileExists)
    {
        $filename = Read-Host $userMessage
    }

    Log "Parsing file: $fileName"
    $XmlDoc = [xml](Get-Content $fileName)

    return $XmlDoc
}

<###################################################################
#  Gets the path of the filename specified from the current 
#  directory. Appends the directoryPath to the current directory if 
#  specified.
###################################################################>
function Global:GetFilePath($filename, $directoryPath='')
{
    $directoryPath = Join-Path $(GetScriptPath) $directoryPath
    $filename = $directoryPath + $filename
    return $filename
}

<###################################################################
#  Gets the common properties xml file.  
###################################################################>
function Global:GetCommonPropertiesXmlFile
{
    return GetXmlFile "CommonProperties.xml" "Enter Common Properties xml filename"
}

<###################################################################
#  Gets the Web Url from the common properties xml file.  
###################################################################>
function Global:GetWebUrl($environment)
{
    $commonXml = GetCommonPropertiesXmlFile
    return GetXMLValue $commonXml "CommonProperties" $environment "/Web/@Url"
}

<###################################################################
#  Gets the Site Url from the common properties xml file.  
###################################################################>
function Global:GetSiteUrl($environment)
{
    $commonXml = GetCommonPropertiesXmlFile
    return GetXMLValue $commonXml "CommonProperties" $environment "/Site/@Url"
}

<###################################################################
#  Gets the Web Application Url from the common properties xml file.  
###################################################################>
function Global:GetWebApplicationUrl($environment)
{
    $commonXml = GetCommonPropertiesXmlFile
    return GetXMLValue $commonXml "CommonProperties" $environment "/WebApplication/@Url"
}

<###################################################################
#  Uses the current script file path to create an xml filename by
#  removing the path and extension.
###################################################################>
function Global:GetXmlFilename($scriptFilepath)
{
    $xmlFilename = ([IO.FileInfo]$scriptFilepath).BaseName + ".xml"
    return $xmlFilename
}

<###################################################################
#  Recycle the SharePoint application pool for the url specified.
###################################################################>
function Global:RecycleSharePointAppPool($siteUrl)
{
    $web = Get-SPWeb $siteUrl
    $sharePointAppPoolName = $web.Site.WebApplication.ApplicationPool.Name

    Log "Recycling Application Pool: $sharePointAppPoolName"

    $serverManager = new-object Microsoft.Web.Administration.ServerManager 
    $serverManager.ApplicationPools | ? { $_.Name -eq $sharePointAppPoolName} | % { $_.Recycle() }

    Log "Recycled Application Pool: $sharePointAppPoolName"
}

<###################################################################
#  Unzips the zip file to the destination specified. If the 
#  destination directory does not exist it is created.
###################################################################>
function Global:Expand-ZipFile($file, $destination)
{
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    
    # Check directory exists - if not create it.
    if(!(Test-Path -Path $destination )){
        New-Item -ItemType directory -Path $destination
        Log "Directory Created: $destination"
    }
    
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
    }
}

<###################################################################
# Get the Search Service Application.
###################################################################>
function Global:GetServiceApplication($searchAppName = $null) {
    $searchServiceApp = $null
    if($searchAppName) {
	    $searchServiceApp = Get-SPEnterpriseSearchServiceApplication -Identity $searchAppName -ErrorAction:SilentlyContinue
	    if($searchServiceApp -eq $null) { TerminateScript "Search Service Application '$searchAppName' could not be found!" }
    } else {
	    $searchServiceApp = Get-SPEnterpriseSearchServiceApplication
	    if($searchServiceApp -eq $null) { TerminateScript "No Search Service Application could not be found!" }
    }

    return $searchServiceApp
}

<###################################################################
# Get the Web Application Url from the Site Url by stripping the 
# end of the url. 
###################################################################>
function Global:GetWebApplicationUrlFromSiteUrl($siteUrl)
{
    # Expect format of siteUrl like 'http://webapplication/css/search' for example.
    $index = $siteUrl.IndexOf("/", 7)

    if ($index -ne -1)
    {
        $webApplicationUrl = $siteUrl.Substring(0, $index)
    }

    return $webApplicationUrl
}

<###################################################################
# Creates a new managed path for the web application if it doesn't
# already exist.
###################################################################>
function Global:CreateManagedPath([string]$webApplicationUrl, [string]$managedPathName, [bool]$explicit=$false)
{
    $managedPaths = Get-SPManagedPath -WebApplication $webApplicationUrl

    if ($managedPaths.Name -contains $managedPathName)
    {
        LogWarning "Managed path '$managedPathName' already exists in web application at '$webApplicationUrl'"
    }
    else
    {
        if ($explicit)
        {
            New-SPManagedPath $managedPathName -WebApplication $webApplicationUrl -Explicit
        }
        else
        {
            New-SPManagedPath $managedPathName -WebApplication $webApplicationUrl
        }
    }
}

<###################################################################
# Removes the crawled property by first unmapping it from any 
# managed properties and then deleting unmapped properties from the 
# category.
###################################################################>
function RemoveCrawledProperty($crawledPropertyName, $categoryName)
{
    $category = Get-SPEnterpriseSearchMetadataCategory -Identity $categoryName -SearchApplication $searchapp
    $crawledProperty = Get-SPEnterpriseSearchMetadataCrawledProperty -Name $crawledPropertyName -SearchApplication $searchapp -Category $category

    if ($crawledProperty)
    {
        $mappings = Get-SPEnterpriseSearchMetadataMapping -SearchApplication $searchapp -CrawledProperty $crawledProperty

        if ($mappings)
        {
            $mappings | Remove-SPEnterpriseSearchMetadataMapping -Confirm:$false
        }
        else
        {
            LogWarning "No mappings found for '$crawledPropertyName'."
        }

        $crawledProperty.IsMappedToContents = $false
        $crawledProperty.Update()
        $category.DeleteUnmappedProperties()

        Log "Deleted crawled property '$crawledPropertyName' from category '$categoryName'"
    }
    else
    {
        LogWarning "Crawled property '$crawledPropertyName' not found."
    }
}

<###################################################################
# Gets the log file with a date added to the filename to allow easy
# identification of when scripts are run.
###################################################################>
function Global:GetLogFilename($scriptFilepath)
{
    $dateTime = Get-Date -format yyyyMMddHHmmss
    $filename = ([IO.FileInfo]$scriptFilepath).BaseName + "." + $dateTime + ".log"
    return $filename
}

<###################################################################
# PowerShell ISE does not support transcript logging.
###################################################################>
function Global:HostSupportsTranscript()
{
    return $Host.Name -ne "Windows PowerShell ISE Host"
}

<###################################################################
# Starts recording all output to a log file.
###################################################################>
function Global:StartLogging($scriptFilepath)
{
    if (HostSupportsTranscript)
    {
        $logFile = GetLogFilename($scriptFilepath)
        $logPath = Join-Path $(GetScriptPath) "\logs\"

        if(!(Test-Path -Path $logPath )){
            New-Item -ItemType directory -Path $logPath
            Log "Directory Created: $logPath"
        }

        $logFile = Join-Path $logPath $logFile
        Start-Transcript -Path $logFile -NoClobber -ErrorAction SilentlyContinue
    }
}

<###################################################################
# Stops recording all output to a log file.
###################################################################>
function Global:StopLogging()
{
    if (HostSupportsTranscript)
    {
        Stop-Transcript
    }
}

<###################################################################
# Checks the file is not null. If it is null it terminates the 
# script.
###################################################################>
function Global:ValidateXMLFile($xml)
{
    if (!$xml)
    {
        TerminateScript "Invalid Xml File - the xml file passed as a parameter is null"
    }
}

<###################################################################
# Gets the environment specific xpath.
###################################################################>
function Global:GetEnvironmentXpath($rootNode, $environment, $xpath)
{
    $xpath = "/$rootNode/Environment/$environment$xpath"
    return $xpath
}

<###################################################################
# Gets the xml value for the specified environment using the xpath
# to find the node selected. If the xpath contains @ it will return 
# an attribute value otherwise it returns the inner text of the node.
###################################################################>
function Global:GetXMLValue($xml, $rootNode, $environment, $xpath)
{
    ValidateXMLFile($xml)

    $xpath = GetEnvironmentXpath $rootNode $environment $xpath
    $xmlNode = $xml.SelectSingleNode($xpath)

    # Determines if its an attribute or node value.
    if ($xpath.Contains("@"))
    {
        $returnValue = $xmlNode.Value
    }
    else
    {
        $returnValue = $xmlNode.InnerText
    }

    return $returnValue 
}

<###################################################################
# Gets the xml nodes for the specified environment using the xpath
# to find the nodes.
###################################################################>
function Global:GetXMLNodes($xml, $rootNode, $environment, $xpath) 
{
    ValidateXMLFile($xml)
    $xpath = GetEnvironmentXpath $rootNode $environment $xpath
    $nodes = $xml.SelectNodes($xpath)
    
    return $nodes
}

<###################################################################
# Sets the file permission on the file/folder specified for the user
# and with the accessLevel specified.
###################################################################>
function Global:SetFileAcl($folderLocation, $username, $accessLevel)
{
    $acl = Get-Acl $folderLocation
    # public FileSystemAccessRule(IdentityReference identity, FileSystemRights fileSystemRights, InheritanceFlags inheritanceFlags, PropagationFlags propagationFlags, AccessControlType type)
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($username, $accessLevel, "ContainerInherit, ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule)
    Set-Acl $folderLocation $acl

    Log "Set file permission on '$folderLocation' to '$accessLevel' for '$username'"
}

<###################################################################
# Deletes the navigation tabs specified from the specified list.
###################################################################>
function Global:DeleteTabs($tabs, $list)
{
    for ($index = $list.ItemCount - 1; $index -ge 0; $index--)
    {
        $item = $list.Items[$index]

        $tab = $tabs | where { $_.Name -eq $item["TabName"].ToString() }
        if ($tab)
        {
            $item.Delete()
            Log "Deleted tab: '$($tab.Name)' from '$($list.Title)'"
        }
    }
}

<###################################################################
# Adds the navigation tab for the page specified to the specified 
# list.
###################################################################>
function Global:AddTab($page, $list)
{
    #Create a new item
    $newItem = $list.Items.Add()
 
    #Add properties to this list item
    $newItem["TabName"] = $page.Title
    $newItem["Page"] = $page.PageName
    $newItem["Tooltip"] = $page.Tooltip
 
    #Update the object so it gets saved to the list
    $newItem.Update()

    Log "Added Search Navigation: '$($searchResultPage.Title)' ($($page.PageName)) to '$($list.Title)'"
}

<###################################################################
# Returns the server relative URL to the xslt file. The xslt file is
# stored in the site collection so get the server relative url of 
# the site.
###################################################################>
function Global:GetServerRelativeXslLink($environment, $xslLink)
{
    $siteUrl = GetSiteUrl($environment)
    $site = Get-SPSite $siteUrl

    if (!$site)
    {
        throw [System.Exception] "Error getting the server relative Url for the XSLLink."
    }

    $serverRelativeUrl = $site.RootWeb.ServerRelativeUrl

    if ($serverRelativeUrl.length -gt 1)
    {
        $xslLink = $serverRelativeUrl + $xslLink
    }

    return $xslLink
}

<###################################################################
# Updates common properties of the webpart if found in the 
# webpartDetail.
###################################################################>
function Global:UpdateWebPartProperties($webPart, $webpartDetail, $environment, $webpartmanager)
{
    $webPartUpdated = $false

    # Update the webpart.
    if ($webPart)
    {                
        # Update the xslt link (Search Core Results webpart).
        if ($webpartDetail.XslLink)
        {
            $webpart.XslLink = GetServerRelativeXslLink $environment $webpartDetail.XslLink 
            $webPartUpdated = $true
            Log "Updating XslLink: $($webpartDetail.XslLink)"
        }

        # Update the location (Search Core Results webpart).
        if ($webpartDetail.Location)
        {
            $webpart.Location = $webpartDetail.Location
            $webPartUpdated = $true
            Log "Updating Location: $($webpartDetail.Location)"
        }

        # Update the Scope (Search Core Results webpart).
        if ($webpartDetail.Scope)
        {
            $webpart.Scope = $webpartDetail.Scope
            $webPartUpdated = $true
            Log "Updating Scope: $($webpartDetail.Scope)"
        }

        # Update the SearchResultPageURL.
        if ($webpartDetail.SearchResultPageURL)
        {
            $webpart.SearchResultPageURL = $webpartDetail.SearchResultPageURL
            $webPartUpdated = $true
            Log "Updating SearchResultPageURL: $($webpartDetail.SearchResultPageURL)"
        }

        # Update the UseLocationVisualization.
        if ($webpartDetail.UseLocationVisualization)
        {            
            $webpart.UseLocationVisualization = [System.Convert]::ToBoolean($webpartDetail.UseLocationVisualization)
            $webPartUpdated = $true
            Log "Updating UseLocationVisualization: $($webpartDetail.UseLocationVisualization)"
        }

        # Update the PropertiesToRetrieve (Search Core Results webpart).
        if ($webpartDetail.PropertiesToRetrieve)
        {            
            $webpart.PropertiesToRetrieve = $webpartDetail.PropertiesToRetrieve
            $webPartUpdated = $true
            Log "Updating PropertiesToRetrieve: $($webpartDetail.PropertiesToRetrieve)"
        }

        # Update the ZoneIndex.
        if ($webpartDetail.ZoneIndex)
        {
            $webpartmanager.MoveWebPart($webPart, $webPart.ZoneID, $webpartDetail.ZoneIndex)
            $webPartUpdated = $true
            Log "Updating ZoneIndex: $($webpartDetail.ZoneIndex)"
        }

        # Update the FilterCategoriesDefinition (Refinement Panel webpart).
        if ($webpartDetail.FilterCategoriesDefinition)
        {
            $webpart.FilterCategoriesDefinition = $webpartDetail.FilterCategoriesDefinition
            $webPartUpdated = $true
            Log "Updating FilterCategoriesDefinition: $($webpartDetail.FilterCategoriesDefinition)"
        }

        if ($webPartUpdated)
        {
            # Save the changes.
            $webpartmanager.SaveChanges($webpart)
            Log "Updated webpart: $($webpartDetail.Title)"
        }
    }

    return $webPartUpdated
}

<###################################################################
# Determines if the permission exists for the specified group on 
# the specified list.
###################################################################>
function Global:DoesListPermissionExist($spList, $groupName, $permissionName)
{
    foreach($roleAssignment in $spList.RoleAssignments)
    {
        if ($roleAssignment.Member.Name -eq $groupName)
        {
            foreach($roleDefinition in $roleAssignment.RoleDefinitionBindings)
            {
                if ($roleDefinition.Type -eq $permission)
                {
                    Log "Permission $($roleDefinition.Name) already exists for list '$($spList.Title)'"
                    return $true
                }
            }
        }
    }

    return $false
}

<###################################################################
# Applies the permission specified to the Group for the list.
###################################################################>
function Global:ApplyListPermission($web, $spList, $groupNameSuffix, $permission)
{
    # Create the group name using the title of the web and the group name.
    $groupName = $web.Title + " " + $groupNameSuffix

    $permissionExists = DoesListPermissionExist $spList $groupName $permission
                
    if (!$permissionExists)
    {        
        # Assign the specified permission to the specified group.
        $spGroup = $web.Groups[$groupName]

        if ($spGroup)
        {
            $assignment = New-Object Microsoft.SharePoint.SPRoleAssignment($spGroup)
            $roleDefinition = $web.RoleDefinitions | Where-Object { $_.Type -eq $permission }

            if ($roleDefinition)
            {
                $assignment.RoleDefinitionBindings.Add($roleDefinition)
                $spList.BreakRoleInheritance($true)
                $spList.RoleAssignments.Add($assignment)
 
                $spList.Update()

                Log "Applied '$permission' Permission to Group '$groupName' on list '$($spList.Title)'"
            }
            else
            {
                Log "Role Definition for '$permission' does not exist in '$($web.Title)' $($web.Url)"
            }
        }
        else
        {
            Log "Group '$groupName' does not exist in '$($web.Title)' $($web.Url)"
        }    
    }
}