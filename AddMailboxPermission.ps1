############################################################################################################################
# OFFICE 365: Add Mailbox Permission (Full Access / SendAs / Auto-mapping)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.3
# Utilizzo:				.\AddMailboxPermission.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		30-04-2014 (06082014)
# Modifiche:
#	0.2- rev1/rev4-correpzioni minori, inclusa adesso la possibilità di modificare ulteriormente le ACL dando anche accesso "Invia Come" (SendAs)
#	0.3- prevedo la possibilità di scegliere l'auto-mapping della casella su Outlook, non utile nel caso di Shared Mailbox che impedirebbero in seguito la ricerca nelle sottocartelle.
############################################################################################################################

#Import-Module MSOnline
""
Write-Host "        Office 365: Add Mailbox Permission" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -foregroundcolor "red"
Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"
$SourceMailbox = Read-Host "Casella alla quale dare accesso (esempio: sharedmailbox@domain.tld)"
$GiveAccessTo = Read-Host "Utente al quale dare Full Access (esempio: mario.rossi@domain.tld) "
""
try
	{
		$title = ""
		$message = "$GiveAccessTo deve caricare automaticamente $SourceMailbox in Outlook (auto-mapping)?"
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$PermissionType = $host.ui.PromptForChoice($title, $message, $options, 0)
		if ($PermissionType -eq 0) {
			Add-MailboxPermission -Identity $SourceMailbox -User $GiveAccessTo -AccessRights FullAccess
			""
			Write-Host "Accesso consentito a $GiveAccessTo (su $SourceMailbox), auto-mapping attivo" -foregroundcolor "green"
			""
			}
		else {		
				""
				Add-MailboxPermission -Identity $SourceMailbox -User $GiveAccessTo -AccessRights FullAccess -AutoMapping:$false
				""
				Write-Host "Accesso consentito a $GiveAccessTo (su $SourceMailbox), auto-mapping DISATTIVATO, ricordarsi di operare sull'Outlook dell'utente per caricare manualmente $SourceMailbox" -foregroundcolor "green"
				""
			}
		
		$title = ""
		$message = "L'utente $GiveAccessTo deve poter inviare a nome di $SourceMailbox ?"
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$PermissionType = $host.ui.PromptForChoice($title, $message, $options, 0)
		if ($PermissionType -eq 0) {
			Add-RecipientPermission $SourceMailbox -Trustee $GiveAccessTo -AccessRights SendAs
			""
			}

		Write-Host "All Done!" -foregroundcolor "green"
		Get-MailboxPermission -Identity $SourceMailbox -User $GiveAccessTo
	}
	catch
	{
		Write-Host "Non riesco ad elaborare l'operazione richiesta. Verificare i dettagli inseriti" -foregroundcolor "red"
		Write-Host $error[0]
		return ""
	}