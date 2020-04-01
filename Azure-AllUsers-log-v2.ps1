# Uninstall-Module AzureAD
# Install-module AzureAD -AllowClobber -Force
# Install-module AzureADPreview -AllowClobber -Force
# Update-Module AzureAD
# 
# Add LogWrite calls along with Write-host calls
 $fdate = (get-date).ToString("yyMM")
 $Logfile = "C:\Scripts\Logs\Azure-AllUSers-Script-v2-$(gc env:computername)-$fdate.log"
# $Logfile = "H:\SharePoint\Admin\Active_Directory\Logs\Azure-AllUSers-Script-v2-$(gc env:computername)$fdate.log"

Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

LogWrite "*****-----------------------------------------------------*****"
LogWrite "***** At the start of the script calling Connect-AzureAD. *****"
#Set up Svc acct credentials
#$password = get-content C:\scripts\cred.txt | convertto-securestring
#$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist "Svc-EOE-It-Apps@mass.gov",$password
#Connect-AzureAd -Credential $credentials

Connect-AzureAd -AccountId "todd.danos@mass.gov"
Clear
# Groups created using the following code.
# New-AzureADGroup -DisplayName "DOE-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "RGT-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "EOE-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "CTF-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "EEC-AllUsers" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"

# New-AzureADGroup -DisplayName "DOE-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "RGT-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "EOE-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "CTF-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
# New-AzureADGroup -DisplayName "EEC-svcRes" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"


# ID's from https://aad.portal.azure.com/#blade/Microsoft_AAD_IAM/GroupsManagementMenuBlade/AllGroups

# AzureAD Groups: EOE-AllUsers, EEC-AllUsers, RGT-AllUsers, DOE-AllUsers, EOE-AllUsers
 
                                       # ID= Agency-AllUsers, 
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
    Try {
         $admembers = Get-AzureADGroupMember -All:$true -ObjectId $group.ID
    } Catch {
        Write-Host "######### CATCH ERROR  $group.ID is not a valid ID #########"
#        $grpID = [string]$group.ID
        LogWrite "######### CATCH ERROR  $grpID is not a valid ID #########"
    }

    # Create a list of Service and Resource accounts to ignore.
     Try {
         $svcRes = Get-AzureADGroupMember -All:$true -ObjectId $group.svcRes
     } Catch {
         Write-Host "######### CATCH ERROR  $group.svcResis not a valid ID #########"
#         $grpSvcRes = [string]$group.svcResis
         LogWrite "######### CATCH ERROR  $grpSvcRes not a valid ID #########"
     }

    # Builid user query: Active and "Department eq Agency"  Removed: or extensionAttribute12 eq Agency

    $query = "accountEnabled eq true and userType eq 'Member' and (Department eq '" + $Group.Agency + "')"   # or extension_4773cea3b62248968250f532a918ce65_extensionAttribute12 eq '" + $Group.Agency + "')"

    # Build List of Users by Agency Department Code
    $users = Get-AzureAdUser -All:$true -Filter $query

    Write-host "***************************"
    Write-host "Agency: " $group.Agency 
    Write-host "    ID: " $group.id
    Write-Host " Count: " $users.count
    $grpAgency = [string]$group.Agency
    $grpID = [string]$group.id
    $usrCount = [string]$users.count
    LogWrite "***************************"
    LogWrite "Agency:  $grpAgency"
    LogWrite "    ID: $grpID"
    LogWrite " Count: $usrCount"

    # Loop through Users
    foreach ($user in $users) {

    # Skip if User is a member of the Group allready.
    if($admembers.ObjectId -contains $user.ObjectId) { 
    # write-host "!!!! Member" $group.Agency " "  $user.DisplayName 
    continue } 

    # Skip if User is a service or Resource account.
    if($svcRes.ObjectId -contains $user.ObjectId) { 
    Write-Host "XXXX svcRes" $group.Agency " "  $user.DisplayName 
    $grpAgency = [string]$group.Agency
    $usrDisplayName = [string]$user.DisplayName
    LogWrite "XXXX svcRes $grpAgency $usrDisplayName"
    continue } 
 
        # Skip service accounts or DL lists. Skip missing Department (excluded from the GAL)
        if ($user.Surname -eq $null -or $user.GivenName -eq $null -or $user.Department -eq $null) { 
        #  Write-Host "#### Skipped" $group.Agency " "  $user.DisplayName 
        continue 
        }

     

          ## Added security to ensure that extensionAttribute12 contains one of the correct agency codes.
               if ($Groups.Agency -contains $user.ExtensionProperty.extension_4773cea3b62248968250f532a918ce65_extensionAttribute12 ) {
        
                Try {      

                        Add-AzureADGroupMember -ObjectId $group.id -RefObjectId $user.ObjectId
                             Write-Host $user.Department " " $user.ExtensionProperty.extension_4773cea3b62248968250f532a918ce65_extensionAttribute12 " " $user.ObjectId " " $user.DisplayName " " $user.Mail
                             $usrDept = [string]$user.Department
                             $usrExtPropExt_4773 = [string]$user.ExtensionProperty.extension_4773cea3b62248968250f532a918ce65_extensionAttribute12
                             $usrObjID = [string]$user.ObjectId
                             $usrDisplayName = [string]$user.DisplayName
                             $usrMail = [string]$user.Mail
                             LogWrite "$user.Department $usrExtPropExt_4773 $usrObjID $usrDisplayName $usrMail"
                                } catch { 
                                 Write-Host "################ CATCH ERROR ################"
                                 Write-Host $user.DisplayName " ********* Error"
                                 $usrDisplayName = [string]$user.DisplayName
                                 LogWrite "################ CATCH ERROR ################"
                                 LogWrite "$usrDisplayName ********* Error"
                                  continue
                         }
                  }

        }
   # Outer Loop
}
