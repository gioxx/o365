$OutputFile = "C:\temp\ClutterDetails.csv"   #The CSV Output file that is created, change for your purposes 
Out-File -FilePath $OutputFile -InputObject "UserPrincipalName,SamAccountName,ClutterEnabled" -Encoding UTF8
 
Write-Host "Retrieving Users"
$objUsers = Get-Mailbox -ResultSize Unlimited | select UserPrincipalName, SamAccountName
 
#Iterate through all users     
Foreach ($objUser in $objUsers) 
{     
    #Prepare UserPrincipalName variable 
    $strUserPrincipalName = $objUser.UserPrincipalName 
    $strSamAccountName = $objUser.SamAccountName
    
    write-host "Processing $strUserPrincipalName"
    #Get Clutter info to the users mailbox 
    $strClutterInfo = $(get-clutter -Identity $($objUser.UserPrincipalName)).isenabled  
    
    #Prepare the user details in CSV format for writing to file 
    $strUserDetails = "$strUserPrincipalName,$strSamAccountName,$strClutterInfo"
     
    #Append the data to file 
    Out-File -FilePath $OutputFile -InputObject $strUserDetails -Encoding UTF8 -append 
} 

Write-Host "Completed - data saved to $OutputFile"