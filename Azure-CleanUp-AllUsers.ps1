# Uninstall-Module AzureAD
# Install-module AzureAD -AllowClobber -Force
# Update-Module AzureAD
# 
# Connect-AzureAd
Clear

<#
# Groups created using the following code.
# New-AzureADGroup -DisplayName "DOE-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "RGT-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "EOE-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "CTF-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "EEC-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"

# Exclude Groups. I.e., Service Accounts, etc.
# New-AzureADGroup -DisplayName "DOE-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "RGT-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "EOE-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "CTF-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "EEC-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"

 ID's from https://aad.portal.azure.com/#blade/Microsoft_AAD_IAM/GroupsManagementMenuBlade/AllGroups

 This script will query all users by Agency Department code then, add the users to one of
 the follow Azure Ad Groups, namely EOE-AllUsers, EEC-AllUsers, RGT-AllUsers, DOE-AllUsers, EOE-AllUsers.
 Note: Any users in corrisponding exlucde Groups will not be added.
 
 #>


 $Groups = @(
  [pscustomobject]@{Agency='CTF';ID='27a77aa4-05a3-4ea8-842b-661d92727866'; svcRes='7164fef4-1aca-42fa-84b4-25c5ba629c12'}
  [pscustomobject]@{Agency='EEC';ID='cb7c0c50-7d6c-486d-9739-0fdb56db174c'; svcRes='e5c8904d-90fe-409f-acfd-705a8b5084dd'}                
  [pscustomobject]@{Agency='RGT';ID='e6cc3441-08a3-46bf-9bdf-087b80db2cc7'; svcRes='262adc5f-21a6-4fbc-a000-cb9b41a8a6c4'}
  [pscustomobject]@{Agency='DOE';ID='c7ccbb8d-eaf2-408d-a072-628eb3a5d0e3'; svcRes='bb078b97-385f-4ab3-9271-d12fc1fcc47c'}
  [pscustomobject]@{Agency='EOE';ID='4a143657-1583-4bfe-9c29-6721ab9b4e45'; svcRes='ddadbb2f-8a75-4d0a-b702-d3414ae450aa'}
   )

# Loop through $Groups list above.
foreach ($group in $groups) {

# Created a user list of existing members in the Group.

    $admembers = Get-AzureADGroupMember -All:$true -ObjectId $group.ID
       $svcRes = Get-AzureADGroupMember -All:$true -ObjectId $group.svcRes

    Write-Host "`nGroup: " $Group.Agency " Members: " $admembers.Count " Service Accounts:" $svcRes.Count "`n"
    # Loop through Members
    foreach ($member in $admembers) {
    
    $Enabled = $member.AccountEnabled
    If (!$Enabled) {  
     Write-Host $member.DisplayName " " $member.AccountEnabled     
    Remove-AzureADGroupMember -ObjectId $group.ID -MemberId $member.ObjectId

    } # End If
    
    } # End Admembers
    
# ####################
    # Loop through Members
    foreach ($member in $svcRes) {
    
    $Enabled = $member.AccountEnabled
    If (!$Enabled) {  
     Write-Host $member.DisplayName " " $member.AccountEnabled     
   Remove-AzureADGroupMember -ObjectId $group.svcRes -MemberId $member.ObjectId

    } # End If
    
    } # End Admembers
    

  } # End Groups
   
  