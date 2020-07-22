<#  https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/SPOSites
# Best practice is to run this script with O365 Global Admin Account.
# SharePoint Admins will have to comment out Line # 75. Only global admins can get accurate reports, i.e., HubSite membership, etc.
#  
# Written by William Gunning 
# 9/10/2019

# Creates a Site Inventnory and then Saves all the Data to a SharePoint Online Custom list.
# V1.1 Added Array for future output to HTML
# v1.2 Added PnP connect for output to SharePoint Online
# v1.3 Tested against a non-GCC tenant.


#>

# Install-Module Microsoft.Online.SharePoint.Powershell
# Import-Module Microsoft.Online.SharePoint.Powershell

# Connection 1 of 2
Write-Host "Admin Credentials to MassGov-admin.SharePoint.com"
 $credentials = Get-Credential;
 Connect-SPOService -URL https://MassGov-admin.sharepoint.com -Credential  $credentials

    # Connection 2 of 2
    Write-Host "Site User Crentials to https://massgov.sharepoint.com/sites/eoe-it-sp"
    $PnPSiteURL = "https://massgov.sharepoint.com/sites/eoe-it-sp"
    Connect-PnPOnline –Url $PnPSiteURL –UseWebLogin

         # Pre-load the Lists into Array
        $ListName = "SPOSites"
        $ListItems  = (Get-PnPlistItem -List $ListName -Field "ID","Title","Url").fieldValues

  $groups = @(
      [pscustomobject]@{Agency='CTF';Path='https://MassGov.sharepoint.com/sites/ctf*'}
      [pscustomobject]@{Agency='EEC';Path='https://MassGov.sharepoint.com/sites/eec*'}                
      [pscustomobject]@{Agency='RGT';Path='https://MassGov.sharepoint.com/sites/rgt*'}
      [pscustomobject]@{Agency='DOE';Path='https://MassGov.sharepoint.com/sites/doe*'}
      [pscustomobject]@{Agency='EOE';Path='https://MassGov.sharepoint.com/sites/eoe*'}
      [pscustomobject]@{Agency='EDU';Path='https://MassGov.sharepoint.com/sites/edu*'}
      )


# https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/get-sposite?view=sharepoint-ps
# Note: Get-SPOSite - Currently, Filter parameter is not functional. Use the Where-Object 


  Foreach ($group in $groups) {
   $sites = Get-SPOSite -Limit All  | Where-Object { $_.Url -like $group.path }
   Write-Host "	" 
   Write-Host "Agency: " $group.agency "Sites: " $sites.count  
   
    $Allinfo = @()  # Reset
    
    Foreach($site in $sites) {           
  
    $extuser = ""; $Users = "";

    try 
    {
        for ($i=0;;$i+=50) 
        {
            $extUser += Get-SPOExternalUser -SiteUrl $site.Url -PageSize 50 -Position $i -ea Stop | Select DisplayName,EMail,AcceptedAs,WhenCreated,InvitedBy | Format-Table | out-string
        }
    }
    catch {
    }

     Write-host processing $ScriptName for $Site.Title $site.Url
    $tld = $Site.URL.Split("/")[3]
    $givenURL = $Site.URL.Split("/")[4]
    $Sitely = $tld + "/" + $givenURL 

                # SharePoint Admins do-not have access to Get-SPOUser command,
                # unless the account has been added to each and every site!
    # $Users =  Get-SPOuser -site $site.Url | Format-Table | out-string
      
         $Properties = @{
            Title = $site.Title
            Sitely = $Sitely
            Owner = $site.Owner
            Agency = $group.agency
            LastContentModifiedDate = $site.LastContentModifiedDate
            Url = $site.Url
            AllowDownloadingNonWebViewableFiles = $site.AllowDownloadingNonWebViewableFiles
            AllowEditing = $site.AllowEditing
            AllowSelfServiceUpgrade = $site.AllowSelfServiceUpgrade
            CommentsOnSitePagesDisabled = $site.CommentsOnSitePagesDisabled
            CompatibilityLevel = $site.CompatibilityLevel
            ConditionalAccessPolicy = $site.ConditionalAccessPolicy
            DefaultLinkPermission = $site.DefaultLinkPermission
            DefaultSharingLinkType = $site.DefaultSharingLinkType
            DenyAddAndCustomizePages = $site.DenyAddAndCustomizePages
            DisableAppViews = $site.DisableAppViews
            DisableCompanyWideSharingLinks = $site.DisableCompanyWideSharingLinks
            DisableFlows = $site.DisableFlows
            DisableSharingForNonOwnersStatus = $site.DisableSharingForNonOwnersStatus
            GivenURL = $givenURL
            GroupID = $site.GroupId
            HubSiteId = $site.HubSiteId
            IsHubSite = $site.IsHubSite
            LimitedAccessFileType = $site.LimitedAccessFileType
            LocaleId = $site.LocaleId
            LockIssue = $site.LockIssue
            LockState = $site.LockState
            PWAEnabled = $site.PWAEnabled
            RelatedgroupId = $site.RelatedGroupId
            ResourceQuota = $site.ResourceQuota
            ResourceQuotaWarningLevel = $site.ResourceQuotaWarningLevel
            ResourceUsageAverage = $site.ResourceUsageAverage
            ResourceUsageCurrent = $site.ResourceUsageCurrent
            RestrictedToGeo = $site.RestrictedToGeo
            SandboxedCodeActivationCapability = $site.SandboxedCodeActivationCapability
            SensitivityLabel = $site.SensitivityLabel
            SharingAllowedDomainList = $site.SharingAllowedDomainList
            SharingBlockedDomainList = $site.SharingBlockedDomainList
            SharingCapability = $site.SharingCapability
            SharingDomainRestrictionMode = $site.SharingDomainRestrictionMode
            ShowPeoplePickerSuggestionsForGuestUsers = $site.ShowPeoplePickerSuggestionsForGuestUsers
            SiteDefinedSharingCapability = $site.SiteDefinedSharingCapability
            SocialBarOnSitePagesDisabled = $site.SocialBarOnSitePagesDisabled
            Status = $site.Status
            StorageQuota = $site.StorageQuota
            StorageQuotaType = $site.StorageQuotaType
            StorageQuotaWarningLevel = $site.StorageQuotaWarningLevel
            StorageUsageCurrent = $site.StorageUsageCurrent
            Template = $site.Template
            TLD = $tld
            WebsCount = $site.WebsCount
            ExternalUser = $extUser
            Users = $users
            
        } # End of Properties
        
        $Allinfo += New-Object psobject -Property $properties
             
         } # End of $sites Foreach
     
        Foreach($Info in $Allinfo){  
        
           # Add or Update Site?                                     # Only Return one Object.
           $ListItem = $ListItems | where {($_.Url -eq $Info.Url)} | Select-Object -First 1
           If ($listitem.Url -eq $Info.Url) { 
           Write-Host "# Update" $Listitem.ID $Info.Title $Info.Url "######" $Adms.MailNickname
          

                Set-PnPListItem -List $ListName -Identity $Listitem.ID -Values @{"Title" = $Info.Title;
                      "Agency" = $info.Agency;
                      "Url" = $info.Url;               
                      "Sitely" = $Info.Sitely;
                      "Owner" = $Info.Owner;
                      "Template" = $Info.Template;
                      "TLD" = $Info.TLD;
                      "WebsCount" = $Info.WebsCount;
                      "ExternalUser" = $Info.ExternalUser;
                      "Users" = $Info.Users;
                      "SharingAllowedDomainList" = $Info.SharingAllowedDomainList;
                      "SharingBlockedDomainList" = $Info.SharingBlockedDomainList;
                      "SharingCapability" = $Info.SharingCapability;
                      "SharingDomainRestrictionMode" = $Info.SharingDomainRestrictionMode;
                      "SiteDefinedSharingCapability" = $Info.SiteDefinedSharingCapability;
                      "GroupID" = $Info.GroupID;
                      "HubSiteId" = $Info.HubSiteId;
                      "IsHubSite" = $Info.IsHubSite;
                      "ConditionalAccessPolicy" = $Info.ConditionalAccessPolicy;
                      "LastContentModifiedDate" = $Info.LastContentModifiedDate;
                      "DisableSharingForNonOwnersStatus" = $Info.DisableSharingForNonOwnersStatus;
                 }                               # End of Set-PnPListItem

                     } ELSE {
                     Write-Host "# Add" $Info.Title $Info.Url
                     Add-PnPListItem -List $ListName -Values @{"Title" = $Info.Title;
                      "Agency" = $info.Agency;
                      "Url" = $info.Url;               
                      "Sitely" = $Info.Sitely;
                      "Owner" = $Info.Owner;
                      "Template" = $Info.Template;
                      "TLD" = $Info.TLD;
                      "WebsCount" = $Info.WebsCount;
                      "ExternalUser" = $Info.ExternalUser;
                      "Users" = $Info.Users;
                      "SharingAllowedDomainList" = $Info.SharingAllowedDomainList;
                      "SharingBlockedDomainList" = $Info.SharingBlockedDomainList;
                      "SharingCapability" = $Info.SharingCapability;
                      "SharingDomainRestrictionMode" = $Info.SharingDomainRestrictionMode;
                      "SiteDefinedSharingCapability" = $Info.SiteDefinedSharingCapability;
                      "GroupID" = $Info.GroupID;
                      "HubSiteId" = $Info.HubSiteId;
                      "IsHubSite" = $Info.IsHubSite;
                      "ConditionalAccessPolicy" = $Info.ConditionalAccessPolicy;
                      "LastContentModifiedDate" = $Info.LastContentModifiedDate;
                      "DisableSharingForNonOwnersStatus" = $Info.DisableSharingForNonOwnersStatus;
                          }                    # End of Add-PnPListItem 

                       # Notes
                      <#
                      "LastContentModifiedDate" = $Info.LastContentModifiedDate;
                      "AllowDownloadingNonWebViewableFiles" = $Info.AllowDownloadingNonWebViewableFiles;
                      "AllowEditing" = $Info.AllowEditing;
                      "AllowSelfServiceUpgrade" = $Info.AllowSelfServiceUpgrade;
                      "CommentsOnSitePagesDisabled" = $Info.CommentsOnSitePagesDisabled;
                      "CompatibilityLevel" = $Info.CompatibilityLevel;
                      "ConditionalAccessPolicy" = $Info.ConditionalAccessPolicy;
                      "DefaultLinkPermission" = $Info.DefaultLinkPermission;
                      "DefaultSharingLinkType" = $Info.DefaultSharingLinkType;
                      "DenyAddAndCustomizePages" = $Info.DenyAddAndCustomizePages;
                      "DisableAppViews" = $Info.DisableAppViews;
                      "DisableCompanyWideSharingLinks" = $Info.DisableCompanyWideSharingLinks;
                      "DisableFlows" = $Info.DisableFlows;
                      "DisableSharingForNonOwnersStatus" = $Info.DisableSharingForNonOwnersStatus;
                      "GivenURL" = $Info.GivenURL;
                      "GroupID" = $Info.GroupID;
                      "HubSiteId" = $Info.HubSiteId;
                      "IsHubSite" = $Info.IsHubSite;
                      "LimitedAccessFileType" = $Info.LimitedAccessFileType;
                      "LocaleId" = $Info.LocaleId;
                      "LockIssue" = $Info.LockIssue;
                      "LockState" = $Info.LockState;
                      "PWAEnabled" = $Info.PWAEnabled;
                      "RelatedgroupId" = $Info.RelatedgroupId;
                      "ResourceQuota" = $Info.ResourceQuota;
                      "ResourceQuotaWarningLevel" = $Info.ResourceQuotaWarningLevel;
                      "ResourceUsageAverage" = $Info.ResourceUsageAverage;
                      "ResourceUsageCurrent" = $Info.ResourceUsageCurrent;
                      "RestrictedToGeo" = $Info.RestrictedToGeo;
                      "SandboxedCodeActivationCapability" = $Info.SandboxedCodeActivationCapability;
                      "SensitivityLabel" = $Info.SensitivityLabel;
                      "SharingAllowedDomainList" = $Info.SharingAllowedDomainList;
                      "SharingBlockedDomainList" = $Info.SharingBlockedDomainList;
                      "SharingCapability" = $Info.SharingCapability;
                      "SharingDomainRestrictionMode" = $Info.SharingDomainRestrictionMode;
                      "ShowPeoplePickerSuggestionsForGuestUsers" = $Info.ShowPeoplePickerSuggestionsForGuestUsers;
                      "SiteDefinedSharingCapability" = $Info.SiteDefinedSharingCapability;
                      "SocialBarOnSitePagesDisabled" = $Info.SocialBarOnSitePagesDisabled;
                      "Status" = $Info.Status;
                      "StorageQuota" = $Info.StorageQuota;
                      "StorageQuotaType" = $Info.StorageQuotaType;
                      "StorageQuotaWarningLevel" = $Info.StorageQuotaWarningLevel;
                      "StorageUsageCurrent" = $Info.StorageUsageCurrent;
                      "Template" = $Info.Template;
                      "TLD" = $Info.TLD;
                      "WebsCount" = $Info.WebsCount;
                      "ExternalUser" = $Info.ExternalUser;
                      "Users" = $Info.Users;
                                              #>

                }                               # End of Else Statement


        }                                       # End of $Allinfo Foreach
       

	}                                           # End of $groups Foreach

