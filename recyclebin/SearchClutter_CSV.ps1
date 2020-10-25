$OutputFile = "C:\temp\clutter\ClutterDetails.csv"   #The CSV Output file that is created, change for your purposes 
Out-File -FilePath $OutputFile -InputObject "UserPrincipalName,SamAccountName,ClutterEnabled" -Encoding UTF8 -Append

Write-Host "Retrieving Users"
Import-Csv C:\temp\MailboxesList.csv | foreach {
	$objUser = get-mailbox $_.UserPrincipalName | select UserPrincipalName, SamAccountName
    $strUserPrincipalName = $objUser.UserPrincipalName 
    $strSamAccountName = $objUser.SamAccountName
	Write-Host "Processing $strUserPrincipalName"
    $strClutterInfo = $(get-clutter -Identity $($objUser.UserPrincipalName)).isenabled  
    $strUserDetails = "$strUserPrincipalName,$strSamAccountName,$strClutterInfo"
    Out-File -FilePath $OutputFile -InputObject $strUserDetails -Encoding UTF8 -append 
} 

Write-Host "Completed - data saved to $OutputFile"