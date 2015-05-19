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
    $serviceContext = Get-SPServiceContext $siteUrl
    $catalog = Get-SPBusinessDataCatalogMetadataObject -BdcObjectType Catalog -ServiceContext $serviceContext
    $models = GetXMLNodes $xmlDoc "BDCProperties" $environment "/Models/Model"

    # Import the models.
    foreach ($model in $models)
    {
        Import-SPBusinessDataCatalogModel -Identity $catalog -Path $($model.Filename) -Force -ErrorAction Stop
        Log "Imported BDC Model '$($model.Name)' (Filename:$($model.Filename))"

        $newModel = Get-SPBusinessDataCatalogMetadataObject -Name $($model.Name) -Namespace $($model.Namespace) -BdcObjectType LobSystemInstance -ServiceContext $serviceContext

        if ($newModel)
        {
            Set-SPBusinessDataCatalogMetadataObject -Identity $newModel -PropertyName "RdbConnection Data Source" -PropertyValue $model.DatabaseServer
        }
        else
        {
            LogError "Cannot find BDC Model '$($model.Name)'"
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