# This script updates the following Custom Lists
#
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/G3
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/Unlicensed

<#
# Written by Bill Gunning on 9-9-2019
# 

     == GCC License Report List of G3 == V3.0

Purpose: Creates a list of All EOE users based on GCC license.

== To use ==
1. Set the variable $SiteURL to the SharePoint site containing the Custom List. 
2. Set the variable $ListName to display name of the custom list "G3"
3. Recommended to index the SPO columns that are being queried or filtered by.

== Required Modules ==
AzureAD
SharePointPnPPowerShellOnline
Install-Module -Name SharePointPnPPowerShellOnline # requires elevation
update-Module -Name SharePointPnPPowerShellOnline # requires elevat

== Triggers: ==
manual excution

== Feature Requests ==
??

== Change log == 
2.0 Changne Title Column to the user's GUID ID
2.1 Added "$UnlicensedItem"
2.1 Fixed issues with deleting unlicensed vs licensed items.
3.0 Added Pre-Load of List Items into Array for Update or Add logic.

== List Views == 
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/G3/AllItems.aspx
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/G3/EOE.aspx
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/G3/CTF.aspx
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/G3/DHE.aspx
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/G3/EEC.aspx
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/G3/EOE.aspx
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/G3/ESE.aspx
#>

# Debug gUser
#$user = Get-Azureaduser -ObjectId ""
#$user = Get-AzureADUser -Filter "userPrincipalName eq 'jondoe@contoso.com'

$credentials = Get-Credential;

$SiteURL = "https://massgov.sharepoint.com/sites/eoe-it-sp"

# MFA Connection
# Connect-AzureAD 
 Connect-PnPOnline –Url $SiteURL –UseWebLogin

# Non-MFA

Connect-AzureAd -Credential $credentials
#Connect-PnPOnline –Url $SiteURL -Credential $credentials

# List Display name, include spaces if required.
$ListName = "G3"
$Unlicensed = "Unlicensed"
 
 $Groups = @(
  [pscustomobject]@{Agency='CTF'}
  [pscustomobject]@{Agency='EEC'}                
  [pscustomobject]@{Agency='RGT'}
  [pscustomobject]@{Agency='DOE'}
  [pscustomobject]@{Agency='EOE'}
   )

 # Pre-load the Lists into Arrays
 $ListItems  = (Get-PnPlistItem -List $ListName -Field "ID","Title","Department").fieldValues
 $UnlicensedItems = (Get-PnPlistItem -List $Unlicensed -Field "ID","Title","Department").fieldValues
 
# Loop through $Groups list above.
foreach ($group in $groups) {

    
    $UserQuery = "accountEnabled eq true and userType eq 'Member' and (Department eq '" + $Group.Agency + "')"  
    $users = Get-AzureAdUser -All:$true -Filter $UserQuery

    Write-host "***************************"
    Write-host "Agency: "  $group.Agency 
    Write-Host " Count: " $users.count       

        # Loop through Users
        foreach ($user in $users) {

       # Write-Host $user.DisplayName $user.ObjectId

        $Rms = ""; $Exchange = ""; $Flow = ""; $Forms =""; $Teams = "";
        $Planner = ""; $ProPlus = ""; $SharePoint = ""; $Skype = "";
        $Stream = ""; $SPOnloin = ""; $PowerApps = "";
        
        $License = Get-AzureADUserLicenseDetail -ObjectId $user.ObjectId
        $RMS = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "RMS_S_ENTERPRISE_GOV"} # Windows Azure Active Directory Rights Management
        $Exchange = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "EXCHANGE_S_ENTERPRISE_GOV"} # Exchange Plan 2G
        $Flow = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "FLOW_O365_P2_GOV"} # Microsoft Flow Plan 2
        $Forms = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "FORMS_GOV_E3"} # Microsoft Forms
        $Teams = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "TEAMS_GOV"}  # Microsoft Teams
        $Planner = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "PROJECTWORKMANAGEMENT_GOV"} # Planner = PROJECTWORKMANAGEMENT
        $ProPlus = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "OFFICESUBSCRIPTION_GOV"} # "Office ProPlus"
        $SharePoint = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "SHAREPOINTENTERPRISE_GOV"} # SharePoint Plan 2G
        $Skype = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "MCOSTANDARD_GOV"} # 	Lync Plan 2G
        $Stream = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "STREAM_O365_E3_GOV"} #Stream
        $SPOnline = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "SHAREPOINTWAC_GOV"} # Office Online for Government
        $PowerApps = $License.ServicePlans | where-Object{$_.ServicePlanName -eq "POWERAPPS_O365_P2_GOV"} # PowerApps

   # If the user has a license
   If ($user.AssignedLicenses) {

# User found. Update List Item
    # Get Item from Array
    $ListItem = $ListItems | where {($_.Title -eq $user.ObjectId)}
    If ($listitem.Title -eq $user.ObjectId) { 
    
    Write-Host "# Update" $user.UserPrincipalName $ListItem.id
    <#
    Set-PnPListItem -List $ListName -Identity $ListItem.id -Values @{"Title" = $user.ObjectId;
   "Department"= $user.Department;
   "UPN"= $user.UserPrincipalName; "User_x0020_Type" = $user.UserType; 
   "RMS"= $RMS.ProvisioningStatus;                   
   "Exchange"= $Exchange.ProvisioningStatus;         
   "Flow"= $Flow.ProvisioningStatus;                 
   "Forms"= $Forms.ProvisioningStatus;                
   "Teams"= $Teams.ProvisioningStatus;             
   "Planner"= $Planner.ProvisioningStatus;         
   "ProPlus"= $ProPlus.ProvisioningStatus;         
   "SharePoint"= $SharePoint.ProvisioningStatus;   
   "Skype"= $Skype.ProvisioningStatus;             
   "Stream"= $Stream.ProvisioningStatus;           
   "Online"= $SPOnline.ProvisioningStatus;         
   "PowerApps"= $PowerApps.ProvisioningStatus;  } #>
   } 
   # AssignedLicenses
     ELSE {         

    # Add new item.
    Write-Host "# Add" $user.UserPrincipalName
   
    Add-PnPListItem -List $ListName -Values @{"Title" = $user.ObjectId;
    "Department"= $user.Department;
    "UPN"= $user.UserPrincipalName; "User_x0020_Type" = $user.UserType; 
    "RMS"= $RMS.ProvisioningStatus;                   #.Equals("Success");
    "Exchange"= $Exchange.ProvisioningStatus;         #.Equals("Success");
    "Flow"= $Flow.ProvisioningStatus;                 #.Equals("Success");
    "Forms"= $Forms.ProvisioningStatus;  
    "Teams"= $Teams.ProvisioningStatus;               #.Equals("Success");
    "Planner"= $Planner.ProvisioningStatus;           #.Equals("Success");
    "ProPlus"= $ProPlus.ProvisioningStatus;           #.Equals("Success");
    "SharePoint"= $SharePoint.ProvisioningStatus;     #.Equals("Success");
    "Skype"= $Skype.ProvisioningStatus;               #.Equals("Success");
    "Stream"= $Stream.ProvisioningStatus;             #.Equals("Success");
    "Online"= $SPOnline.ProvisioningStatus;           #.Equals("Success");
    "PowerApps"= $PowerApps.ProvisioningStatus        #.Equals("Success");
     } 
    
    }

 # Remove  user from Unlicensed List as the user  has a license.
 
 # Get Item from Array
 $UnlicensedItem =  $UnlicensedItems | where {($_.Title -eq $user.ObjectId)}
  If ($UnlicensedItem.Title -eq $user.ObjectId) {
  Write-Host "UnlicensedItems"
 # If Query is not Null, then delete
 Write-Host "-Remove from Unlicensed " $user.UserPrincipalName
# Remove-PnPListItem  -List $Unlicensed -Identity $UnlicensedItem.id -Force
 }

    # End loop AssignedLicenses

    }

     ELSE {
# No Assigned License

# Remove  User from G3 Licensed as the user does not have a license.
    # Get Item from Array
    $ListItem = $ListItems | where {($_.Title -eq $user.ObjectId)}
    If ($listitem.Title -eq $user.ObjectId) { 
 Write-Host "-Remove from G3 " $user.UserPrincipalName
# Remove-PnPListItem  -List $ListName -Identity $ListItem.id -Force
 }

 
 # User found. Update List Item
$UnlicensedItem =  $UnlicensedItems | where {($_.Title -eq $user.ObjectId)}
  If ($UnlicensedItem.Title -eq $user.ObjectId) {
   Write-Host "Unlicesned Update" $user.UserPrincipalName;
   <#
      Set-PnPListItem -List $Unlicensed -Identity $UnlicensedItem.id -Values @{"Title" = $user.ObjectId;
    "Department"= $user.Department;
     "UPN"= $user.UserPrincipalName; "UserType" = $user.UserType} #>
    } else {
 # User not found. Add user to list.
 
 Write-Host "# Add Unlicensed " $user.UserPrincipalName

  Add-PnPListItem -List $Unlicensed  -Values @{"Title" = $user.ObjectId;
    "Department"= $user.Department;
    "UPN"= $user.UserPrincipalName;
    "UserType" = $user.UserType} 
    }

# End loop Unlicensed Licenses


}

   }
   #Group

  }

