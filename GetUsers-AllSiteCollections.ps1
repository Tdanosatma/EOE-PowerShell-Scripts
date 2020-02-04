#################################################################
# Script that allows to get all the users for all the Site Collections in a SharePoint Online Tenant
# Required Parameters:
#  -> $sUserName: User Name to connect to the SharePoint Admin Center.
#  -> $sMessage: Message to show in the user credentials prompt.
#  -> $sSPOAdminCenterUrl: SharePoint Admin Center Url

##################################################################

# Add LogWrite calls along with Write-host calls
$fdate = (get-date).ToString("yyMM")
$Logfile = "C:\temp\SharePoint Online\Users-$(gc env:computername)-$fdate.log"

Function LogWrite
{
    Param ([string]$logstring)
    Add-content $Logfile -value $logstring
}

$host.Runspace.ThreadOptions = "ReuseThread"

#Definition of the function that gets all the site collections information in a SharePoint Online tenant
function Get-SPOUsersAllSiteCollections
{
    param ($sUserName,$sMessage)
    try
    { 
        Write-Host "----------------------------------------------------------------------------" -foregroundcolor Green
        Write-Host "Getting the information for all the site colletions in the Office 365 tenant" -foregroundcolor Green
        Write-Host "----------------------------------------------------------------------------" -foregroundcolor Green
        #   LogWrite "----------------------------------------------------------------------------"
        #   LogWrite "Getting the information for all the site colletions in the Office 365 tenant"
        #   LogWrite "----------------------------------------------------------------------------"
        echo (get-date).ToString("yyyyMMdd") | Out-File $Logfile -Append;
        echo "----------------------------------------------------------------------------" | Out-File $Logfile -Append;
        echo "Getting the information for all the site colletions in the Office 365 tenant" | Out-File $Logfile -Append;
        echo "----------------------------------------------------------------------------" | Out-File $Logfile -Append;

        $msolcred = get-credential -UserName $sUserName -Message $sMessage
        Connect-SPOService -Url $sSPOAdminCenterUrl -Credential $msolcred
        $spoSites=Get-SPOSite | Select *

        foreach($spoSite in $spoSites)
        {
            Write-Host "Users for " $spoSite.Url -foregroundcolor Blue
            $Url = $spoSite.Url
         #   LogWrite "Users for $Url`r`n"
            echo "Users for $Url" | Out-File $Logfile -Append;
         #   Get-SPOUser -Site $spoSite.Url
            Get-SPOUser -Site $spoSite.Url | Out-File $Logfile -Append;
         #   $SPOLUser = Get-SPOUser -Site $spoSite.Url
            Write-Host
         #   LogWrite "`r`n"
            echo "`r`n" | Out-File $Logfile -Append;
        } 
     #   Write-Host "Getting users" -ForegroundColor Green
     #   LogWrite "`r`nGetting users"
     #   Get-SPOUser -Site "https://masseoe.sharepoint.com" | Out-File $Logfile -Append;
    }
    catch [System.Exception]
    {
        write-host -f red $_.Exception.ToString() 
    } 
}

#Connection to Office 365
$sUserName="tdanos@masseoe.onmicrosoft.com"
$sMessage="SPO Credential Please"
$sSPOAdminCenterUrl="https://masseoe-admin.sharepoint.com"
#Get-SPOUser -Site "https://<Domain>.sharepoint.com/" -LoginName "<user>"

Get-SPOUsersAllSiteCollections -sUserName $sUserName -sMessage $sMessage