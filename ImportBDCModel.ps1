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

        $lobSystem = Get-SPBusinessDataCatalogMetadataObject -BdcObjectType LobSystem -Name $($model.LobSystemName) -ServiceContext $serviceContext

        if ($lobSystem)
        {
            $instance = $instance = $lobSystem.LobSystemInstances | Where-Object { $_.Name -eq $model.LobSystemInstanceName }

            if ($instance)
            {
                Set-SPBusinessDataCatalogMetadataObject -Identity $instance -PropertyName "RdbConnection Data Source" -PropertyValue $model.DatabaseServer
            }
            else
            {
                LogError "LobSystemInstance '$($model.LobSystemInstanceName)' not found"
            }
        }
        else
        {
            LogError "LobSystem '$($model.LobSystemName)' not found"
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