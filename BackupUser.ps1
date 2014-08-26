############################################################################################################################
# OFFICE 365: User Backup Operations (Ready2Remove)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\BackupUser.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		06-05-2014 (26082014)
# Modifiche:			
#	0.2- inclusa la richiesta dell'utente a cui dare le autorizzazioni in full-access al database del vecchio dipendente
############################################################################################################################

#Import-Module MSOnline
""
Write-Host "        Office 365: Backup User" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -foregroundcolor "red"
Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"
$ResetUser = Read-Host "Utente (esempio: info@domain.tld)                                 "
$AdminUser = Read-Host "A chi devo dare il Full Access? (esempio: mario.rossi@domain.tld) "
""
try
	{
		Set-MsolUserPassword -UserPrincipalName $ResetUser -NewPassword "Office2013"
		Set-MsolUser -UserPrincipalName $ResetUser -PasswordNeverExpires $true
		#DEBUG Get-MSOLUser -UserPrincipalName $ResetUser | Select PasswordNeverExpires
		Write-Host "$ResetUser: Reset password OK"
		Add-MailboxPermission -Identity $ResetUser -User $AdminUser -AccessRights FullAccess
		Write-Host "$ResetUser: ACL modificate, richiedo dettaglio ..."
		Get-MailboxPermission -Identity $ResetUser -User $AdminUser
		Write-Host ""
		Write-Host "-------------------------------------------------------------"
		Write-Host "Procedere con le operazioni di backup su PST ed eliminazione utente da Exchange"
	}
	catch
	{
		Write-Host "Non riesco a eseguire l'operazione richiesta, riprovare." -foregroundcolor "red"
		write-host $error[0]
		return ""
	}