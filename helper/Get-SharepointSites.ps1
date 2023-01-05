# Uncomment to see Verbose statements, set to "SilentyContinue" to hide them
# $VerbosePreference = "Continue" 
#Requires -PSEdition Core

function Get-SharepointSites {
    $sites = Get-PnPTenantSite
    return $sites
    
}

