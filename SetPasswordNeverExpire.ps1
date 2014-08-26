############################################################################################################################
# OFFICE 365: Set Password to Never Expire (single user)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetPasswordNeverExpire.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		29-05-2014
# Modifiche:			-
############################################################################################################################

#$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $PWord
#Import-Module MSOnline
#Connect-MsolService -Credential $Credential

#Main
Function Main {

	""
	Write-Host "        Office 365: Set Password to Never Expire (single user)" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$RicercaUser = Read-Host "Utente da modificare (esempio: mario.rossi@domain.tld) "
	
	try
	{
		Set-MsolUser -UserPrincipalName $RicercaUser -PasswordNeverExpires $true
		""
		Write-Host "Utente modificato, verifico:" -foregroundcolor "green"
		Get-MSOLUser -UserPrincipalName $RicercaUser | Select UserPrincipalName, PasswordNeverExpires
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
		write-host $error[0]
		return ""
	}
	
}

# Start script
. Main