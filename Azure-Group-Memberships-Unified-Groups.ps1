# Updates the Custom List: UnifiedGroups
# https://massgov.sharepoint.com/sites/EOE-it-sp/Lists/UnifiedGroups

# Uninstall-Module AzureAD
# Install-module AzureAD -AllowClobber -Force
# Update-Module AzureAD

 #$credentials = Get-Credential;

# Connection 1 of 2
#Connect-AzureAd -Credential $credentials

  # Connection 2 of 2
    Write-Host "Site User Crentials to https://massgov.sharepoint.com/sites/eoe-it-sp"
    $PnPSiteURL = "https://massgov.sharepoint.com/sites/eoe-it-sp"
     Connect-PnPOnline –Url $PnPSiteURL –UseWebLogin
  #  Connect-PnPOnline –Url $PnPSiteURL -Credential $credentials

       # Pre-load the Lists into Array
       $ListName = "Unified-Groups"
       $ListItems  = (Get-PnPlistItem -List $ListName -Field "ID","Title").fieldValues

# 
 $Agencies = @(
  [pscustomobject]@{Agency='CTF'}
  [pscustomobject]@{Agency='EEC'}                
  [pscustomobject]@{Agency='RGT'}
  [pscustomobject]@{Agency='DOE'}
  [pscustomobject]@{Agency='EOE'}
 # [pscustomobject]@{Agency='EDU'}
   )

# Loop through each Agency
ForEach ($Agency in $Agencies) {

     $agencyQuery = "startswith(Mail,'" + $Agency.Agency +  "')"
     $Groups = (Get-AzureADGroup -Filter $agencyQuery)   
     Write-Host $Agency.Agency $Groups.count

     # Loop through all Groups until you find a Unified.i
     
     ForEach ($Group in $Groups) {
 
     $groupTypes = "Null"
     Try {
     $adms  = Get-AzureADMSGroup -Id $group.ObjectId
     $groupTypes = $adms.GroupTypes
     } catch { Write-Host "Error" }

     $groupTypes = $adms.GroupTypes
 
     IF ($groupTypes -eq "Unified") {
     
      # Add or Update Site?  
                                                         # Only Return the 1st Object.
           $ListItem = $ListItems | where {($_.Title -eq $Adms.DisplayName) } # | Select-Object -First 1
           
              
           If ($adms.MailNickname -eq $Null) { $MailNickname = "Null"
                    } else { $MailNickname = $adms.MailNickname}

           $Owners = ""; $Owners = Get-AzureADGroupOwner -ObjectId $adms.Id
           $Owner = $Owners.UserPrincipalName -Join "; "
           $Members = ""; $Members = Get-AzureADGroupMember -ObjectId $adms.Id
           $Member = $Members.UserPrincipalName -Join "; "
           
         #  Start-Sleep -Seconds .5

           If ($ListItem) { 
           
           Write-host "# Update" $ListItem.ID $Adms.MailNickname
           
         <#  Set-PnPListItem -List $ListName -Identity $ListItem.ID -Values @{"Title" = $adms.DisplayName;
           "Agency" = $Agency.Agency
           "Description" = $adms.Description;
           "Mail" = $adms.Mail;
           "MailNickname" = $adms.MailNickname;
           "Owners" = $Owner;
           "Members" = $Member;
           "Visibility" = $adms.Visibility
           "SecurityEnabled" = $adms.SecurityEnabled;
           "CreatedDateTime" = $adms.CreatedDateTime;
           "RenewedDateTime" = $adms.RenewedDateTime;
              }  #>
           } Else {
           Write-Host "# Add" $Adms.MailNickname
          Add-PnPListItem -List $ListName -Values @{"Title" = $adms.DisplayName;
           "Agency" = $Agency.Agency
           "Description" = $adms.Description;
           "Mail" = $adms.Mail;
           "MailNickname" = $adms.MailNickname;
           "Owners" = $Owner;
           "Members" = $Member;
           "Visibility" = $adms.Visibility
           "SecurityEnabled" = $adms.SecurityEnabled;
           "CreatedDateTime" = $adms.CreatedDateTime;
           "RenewedDateTime" = $adms.RenewedDateTime;
             }
           } # End else
     

           } # End If $adms.GroupTypes
           
    
 
    }  # End ForEach $Groups
     } # End For Each $Agencies

 