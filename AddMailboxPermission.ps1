############################################################################################################################
# OFFICE 365: Add Mailbox Permission (Full Access / SendAs / GrantSendOnBehalfTo / Auto-mapping)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.7
# Utilizzo:				.\AddMailboxPermission.ps1
#						(opzionale, passaggio dati da prompt) .\AddMailboxPermission.ps1 -SourceMailbox shared@contoso.com -GiveAccessTo mario.rossi@contoso.com
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		06-10-2015
# Modifiche:
#	0.7- Lo script accetta adesso i parametri passati da riga di comando (-SourceMailbox e -GiveAccessTo)
#	0.6- correzioni minori. Messo meglio in evidenza i dettagli riguardanti il "Send As" e il "Send on Behalf to".
#	0.5- modificata la richiesta di GrantSendOnBehalfTo che ora viene mostrata solo se si rifiuta il SendAs.
#	0.4- aggiungo la possibilità di specificare se l'utente deve inviare con proprietà Grantsendonbehalfto e non SendAs completo.
#	0.3- prevedo la possibilità di scegliere l'auto-mapping della casella su Outlook, non utile nel caso di Shared Mailbox che impedirebbero in seguito la ricerca nelle sottocartelle.
#	0.2- rev1/rev4-correpzioni minori, inclusa adesso la possibilità di modificare ulteriormente le ACL dando anche accesso "Invia Come" (SendAs).
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $SourceMailbox, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $GiveAccessTo 
)

""
Write-Host "        Office 365: Add Mailbox Permission" -foregroundcolor "green"
Write-Host "        ------------------------------------------"

if (([string]::IsNullOrEmpty($SourceMailbox) -eq $true) -and ([string]::IsNullOrEmpty($RemoveAccessTo) -eq $true))
{
	#MANCANO I DETTAGLI DA PROMPT, LI RICHIEDO A VIDEO
	
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	""
	$SourceMailbox = Read-Host "Casella alla quale dare accesso (esempio: sharedmailbox@domain.tld)"
	$GiveAccessTo = Read-Host "Utente al quale dare Full Access (esempio: mario.rossi@domain.tld) "
	""
}
else
{
	#DETTAGLI PASSATI DA PROMPT, POSSO PROCEDERE IMMEDIATAMENTE

	try
		{
			""
			Write-Host "Autorizzo $GiveAccessTo a utilizzare $SourceMailbox ..." -f "yellow"
			""
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
					Write-Host "Accesso consentito a $GiveAccessTo (su $SourceMailbox), auto-mapping DISATTIVATO"
					Write-Host "Ricordarsi di operare sull'Outlook dell'utente per caricare manualmente $SourceMailbox" -foregroundcolor "green"
					""
				}
			
			$title = "SendAs - L'utente $GiveAccessTo deve poter inviare come fosse $SourceMailbox ?"
			$message = "ATTENZIONE: questo permette a $GiveAccessTo di scrivere a tutti gli effetti come fosse $SourceMailbox"
			$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
			$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
			$PermissionType = $host.ui.PromptForChoice($title, $message, $options, 0)
			if ($PermissionType -eq 0) {
				Add-RecipientPermission $SourceMailbox -Trustee $GiveAccessTo -AccessRights SendAs
				""
				}
			else {
			
					""
					Write-Host "SendAs non impostato." -f Red
					$title = " "
					$message = "Send On Behalf To - L'utente $GiveAccessTo deve poter almeno inviare a nome di $SourceMailbox (Es. Mario Rossi per conto di $SourceMailbox)?"
					$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
					$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
					$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
					$PermissionType = $host.ui.PromptForChoice($title, $message, $options, 0)
					if ($PermissionType -eq 0) {
						Set-Mailbox $SourceMailbox -GrantSendOnBehalfTo @{add="$GiveAccessTo"}
						""
						}
					else {
						""
						Write-Host "Send On Behalf To non impostato." -f Red
						}
			}

			""
			Write-Host "All Done!" -foregroundcolor "green"
			Get-MailboxPermission -Identity $SourceMailbox -User $GiveAccessTo
		}
		catch
		{
			Write-Host "Non riesco ad elaborare l'operazione richiesta. Verificare i dettagli inseriti" -foregroundcolor "red"
			Write-Host $error[0]
			return ""
		}
}