
<# Provision multiple URLs at once #>
#$URLs = @("https://massgov.sharepoint.com/sites/eoe",
# "https://massgov.sharepoint.com/sites/eoe-it-apps")
#$URLs = @("https://massgov.sharepoint.com/sites/doe-cacn",
#          "https://massgov.sharepoint.com/sites/doe-ssoslt",
#          "https://massgov.sharepoint.com/sites/eoe-it-apps-dev")
 
 Clear
 Write-Host "### Provisioning Content Panda ###"
 Write-Host

<# Provision a single URL #>
 $URLs = "https://massgov.sharepoint.com/sites/doe-seis"

$clientSideComponentId = "af656294-ff42-406d-b282-b2d5d2b9a801"

#Prompt user for tenant admin credentials
$credentials = Get-Credential;

# Production
$clientSideComponentProperties = '{"accountID":"DC54004E-EE5F-498F-A4B4-34FC3B66DEA7"}'

# EOTSS Development Dev tenant accountID:
# $clientSideComponentProperties = '{"accountID":"B9C957C5-F897-4F68-936F-86DB1559B38A"}'

Write-Host

# Loop through site collections and install Content Panda App Extension
foreach ($URL in $URLs)
{
   Connect-PnPOnline -Url $URL -Credentials ($credentials)
   Write-Host "Running installation for:" $URL
  # Write-Host "--Looking for existing action"
   $ContentPandaAction = (Get-PnPCustomAction).Where({ $_.Name -eq "ContentPanda" })
  
# Remove Content Panda from Site. Disable or remove for new upgrade.
        if($ContentPandaAction) {
            Write-Host "--Removing existing action"
            Remove-PnPCustomAction -Identity $ContentPandaAction.Id -Force
             }

# Add Content Panda back to site.

#TJD 
    Write-Host "--Adding Content Panda Action"

## Comment out the line below to remove Content Panda
#TJD     
    Add-PnPCustomAction -Title "ContentPanda" -Name "ContentPanda" -Location "ClientSideExtension.ApplicationCustomizer" -ClientSideComponentId $clientSideComponentId -ClientSideComponentProperties $clientSideComponentProperties

# $URLs End of Loop
}

Write-Host
Write-Host "Installation Complete"
