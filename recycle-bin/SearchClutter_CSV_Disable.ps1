Import-Csv C:\temp\DisableClutter.csv | foreach {
	$objUser = get-mailbox $_.UserPrincipalName | select UserPrincipalName
    $strUserPrincipalName = $objUser.UserPrincipalName 
    write-host "Processing $strUserPrincipalName" 
	Set-Clutter -Identity $strUserPrincipalName -Enable:$False
}