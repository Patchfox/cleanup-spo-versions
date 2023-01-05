

[CmdletBinding()]


$globalConfig = Get-Content -Path $PSScriptRoot\globalConfig.jsonc | ConvertFrom-Json

#Register-PnPAzureADApp -ApplicationName "PnP Evaluation" -Tenant $globalConfig.TenantURL -OutPath c:\mycertificates -DeviceLogin

#Load cmdlets
try {
    . "$PSScriptRoot/helper/Get-SharepointSiteStorage.ps1"
    . "$PSScriptRoot/helper/Get-SharepointSites.ps1"
    . "$PSScriptRoot/helper/Remove-OlderVersions.ps1"
} catch {
   
   
    Write-Host 'Error while loading supporting PowerShell Scripts' 
}

#TODO Authentication/Authorization
#Connect to PnP Online
try {
    Connect-PnPOnline -Url 'https://$(globalConfig.TenantURL}' -ClientId $globalConfig.ClientID -Tenant $globalConfig.TenantURL -CertificateBase64Encoded $globalConfig.Base64EncodedCert

    $AllSpoSites = Get-SharepointSites
    $SPSites = Get-SharepointSiteStorage -storagethreshold $globalConfig.storagethreshold -sites $AllSpoSites
    foreach ($SPsite in $SPSites) {
        Remove-OlderVersions -VersionsToKeep $globalConfig.VersionsToKeep -SiteUrl $SPsite.SiteUrl 
    }

} catch {
    Write-Error $('Fehler aufgetreten:' + $_.Exception.Message) 
}
