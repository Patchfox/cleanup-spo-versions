# Uncomment to see Verbose statements, set to "SilentyContinue" to hide them
# $VerbosePreference = "Continue" 
#Requires -PSEdition Core

function Get-SharepointSiteStorage {

# Uncomment to see Verbose statements, set to "SilentyContinue" to hide them
# $VerbosePreference = "Continue" 

  [CmdletBinding()]
  param (
 
    [string]$storagethreshold,
    [object]$sites
    
  )

  $results = @()

  #Get all sites and calculates the storage consumption of them
  foreach ($site in $sites) {
    $siteStorage = New-Object PSObject
    if (!$site.StorageUsageCurrent -eq 0 -and !$site.StorageQuota -eq 0) {
      $percent = $site.StorageUsageCurrent / $site.StorageQuota * 100
    } else {
      $percent = 0
    }

    $percentage = [math]::Round($percent, 2)
    
    $siteStorage | Add-Member -MemberType NoteProperty -Name 'SiteTitle' -Value $site.Title
    $siteStorage | Add-Member -MemberType NoteProperty -Name 'SiteUrl' -Value $site.Url
    $siteStorage | Add-Member -MemberType NoteProperty -Name 'PercentageUsed' -Value $percentage
    $siteStorage | Add-Member -MemberType NoteProperty -Name 'StorageUsed(MB)' -Value $site.StorageUsageCurrent
    $siteStorage | Add-Member -MemberType NoteProperty -Name 'StorageQuota(MB)' -Value $site.StorageQuota
    
    if ($site.StorageUsageCurrent -gt $storagethreshold ) {
      $results += $siteStorage 
    
    }
    $siteStorage = $null
  }
  return $results
}