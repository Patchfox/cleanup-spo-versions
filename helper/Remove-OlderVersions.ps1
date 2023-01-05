# Uncomment to see Verbose statements, set to "SilentyContinue" to hide them
# $VerbosePreference = "Continue" 
#Requires -PSEdition Core

function Remove-OlderVersions {

    


    [CmdletBinding()]
    param (

        [string]$VersionsToKeep,
        [string]$SiteUrl

    
    )
    Try {
        $globalConfig = Get-Content -Path $PSScriptRoot\..\globalConfig.jsonc | ConvertFrom-Json
        Connect-PnPOnline -Url $SiteUrl -ClientId $globalConfig.ClientID -Tenant $globalConfig.TenantURL -CertificateBase64Encoded $globalConfig.Base64EncodedCert

        #Get the Context
        $Ctx = Get-PnPContext
 
        #Exclude certain libraries
        $ExcludedLists = @('Converted Forms', 'Customized Reports', 'Form Templates', 'Bilder', 'Workflowaufgaben' , 
            'Bilder der Websitesammlung', 'Formatbibliothek', 'List Template Gallery', 'Master Page Gallery', 'Pages', 
            'Reporting Templates', 'Site Assets', 'Aktuelle Liste der Site Websiteobjekte', 'Site Collection Documents', 'Site Collection Images', 'Site Pages', 
            'Solution Gallery', 'Style Library', 'Theme Gallery', 'Web Part Gallery', 'wfpub', 'Inhalts- und Strukturberichte', 
            'Formularvorlagen', 'Websiteobjekte', 'Registerkarten in Suchergebnissen', 'Wiederverwendbarer Inhalt', 'Registerkarten auf Suchseiten' ) 
        Get-PnPList
        #Get All document libraries
        $DocumentLibraries = Get-PnPList | Where-Object {$_.BaseType -eq "DocumentLibrary" -and $_.Title -notin $ExcludedLists -and $_.Hidden -eq $false}
 
        #Iterate through each document library
        ForEach ($Library in $DocumentLibraries) {
            Write-Host 'Processing Document Library:'$Library.Title -f Magenta
 
            #Get All Items from the List - Exclude 'Folder' List Items
            $ListItems = Get-PnPListItem -List $Library -PageSize 2000 | Where-Object { $_.FileSystemObjectType -eq 'File' }
 
            #Loop through each file
            ForEach ($Item in $ListItems) {
                #Get File Versions
                $File = $Item.File
                $Versions = $File.Versions
                $Ctx.Load($File)
                $Ctx.Load($Versions)
                $Ctx.ExecuteQuery()
  
                Write-Host -f Yellow "`t Scanning File:"$File.Name
                $VersionsCount = $Versions.Count
                $VersionsToDelete = $VersionsCount - $VersionsToKeep
                If ($VersionsToDelete -gt 0) {
                    Write-Host -f Cyan "`t Total Number of Versions of the File:" $VersionsCount
                    #Delete versions
                    For ($i = 0; $i -lt $VersionsToDelete; $i++) {
                        $Versions[0].DeleteObject()
                        Write-Host -f Cyan "`t`t Deleted Version:" $Versions[0].VersionLabel
                    }
                    $Ctx.ExecuteQuery()
                    Write-Host -f Green "`t Version History is cleaned for the File:"$File.Name
                }
            }
        }
    } Catch {
        Write-Host -f Red 'Error Cleaning up Version History!' $_.Exception.Message
    }

}