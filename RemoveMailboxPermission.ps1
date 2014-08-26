############################################################################################################################
# OFFICE 365: Remove Mailbox Permission
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\RemoveMailboxPermission.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		25-07-2014
# Modifiche:
#	0.2- Inclusa la rimozione dell'ACL SendAs (prima veniva effettuata la rimozione del solo FullAccess, rimaneva attivo il SendAs)
############################################################################################################################

#Import-Module MSOnline
""
Write-Host "        Office 365: Remove Mailbox Permission" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -foregroundcolor "red"
Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"
$SourceMailbox = Read-Host "Casella da modificare (esempio: sharedmailbox@domain.tld)            "
$GiveAccessTo = Read-Host "Utente al quale togliere l'accesso (esempio: mario.rossi@domain.tld) "
""
try
	{
		""
		Remove-MailboxPermission -Identity $SourceMailbox -User $GiveAccessTo -AccessRights FullAccess -InheritanceType All
		Write-Host "Accesso rimosso per $GiveAccessTo (su $SourceMailbox)" -foregroundcolor "green"
		""
		Write-Host "Tento rimozione parametro SendAs per $GiveAccessTo (su $SourceMailbox)" -foregroundcolor "yellow"
		Remove-RecipientPermission $SourceMailbox -Trustee $GiveAccessTo -AccessRights SendAs
		""
		Write-Host "All Done!" -foregroundcolor "green"
		Write-Host "Riepilogo accessi alla casella di $SourceMailbox " -foregroundcolor "yellow"
		Get-MailboxPermission -Identity $SourceMailbox
	}
	catch
	{
		Write-Host "Non riesco ad elaborare l'operazione richiesta. Verificare i dettagli inseriti" -foregroundcolor "red"
		Write-Host $error[0]
		return ""
	}