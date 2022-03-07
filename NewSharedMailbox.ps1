<#
OFFICE 365: Create New Shared Mailbox
----------------------------------------------------------------------------------------------------------------
Autore originale: 		Alan Byrne
Versione originale:		1.0
Modifiche:						GSolone
Versione:							0.9
Utilizzo:							.\New-SharedMailbox.ps1
Info:									https://gioxx.org/tag/o365-powershell
Ultima modifica:			15-10-2020
Modifiche:
		0.9- imposto MessageCopyForSendOnBehalfEnabled e MessageCopyForSentAsEnabled entrambi a $True per salvare copia delle email inviate dalla casella di posta condivisa direttamente al suo interno e non solo nella Sent Items dell'utente che la sta utilizzando.
		0.8- ho modificato l'esempio relativo all'indirizzo di posta da creare (ora info@contoso.com). Ho messo a posto parte dell'indentazione, e ho aggiunto un blocco che mostra i permessi applicati alla casella di posta quando si termina di aggiungerne.
		0.7- ho commentato la forzatura del MicrosoftOnlineServicesID, non più necessaria
		0.6- ho dovuto inserire un nuovo "sleep" prima di settare il MicrosoftOnlineServicesID perché in alcuni casi il server Exchange non trova immediatamente l'utente, cosa che invece succede già a 10 secondi di distanza dal comando di creazione casella.
		0.5- correzioni minori
		0.1/0-4- rimosso limite Shared a 5GB, forzata la richiesta dell'alias da utilizzare, rimossa richiesta permission Send As e Automap (entrambe dati per scontati e impostati di default), forzato il MicrosoftOnlineServicesID sull'indirizzo completo di posta elettronica
#>

#Main
Function Main {
	""; Write-Host "        Office 365: New Shared Mailbox" -f "Green";
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -f "Red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -f "Red"
	Write-Host "          nei dati richiesti qui di seguito" -f "Red"; "";
	Write-Host "-------------------------------------------------------------------------------------------------"
	$SharedMailboxUserName = Read-Host "Indirizzo (esempio: info@contoso.com)       "
	$SharedMailboxDisplayName = Read-Host "Nome casella (esempio: Contoso srl -Info)   "
	$SharedMailboxAlias = Read-Host "Alias (esempio: Contososrl_info)            "

	try	{
		""; Write-Host "Creo la Shared Mailbox, attendi." -f "Yellow"; ""
		New-Mailbox -Name $SharedMailboxDisplayName -Alias $SharedMailboxAlias -Shared -PrimarySMTPAddress $SharedMailboxUserName
		Start-Sleep -s 10
	} catch	{
		Write-Host "Non riesco a creare la Shared Mailbox, riprovare." -f "Red"
		write-host $error[0]
		return ""
	}

	""; Write-Host "Shared Mailbox creata correttamente: $SharedMailboxUserName" -f "Green"; ""

	#Prompt user to add permissions to mailbox
	$title = "PERMESSI"
	$message = "Vuoi aggiungere adesso i permessi di accesso alla casella?"
	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Aggiungi permessi."
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
	$result = $host.ui.PromptForChoice($title, $message, $options, 0)

	while ($result -eq 0) {
		""; "";
		#User to add to mailbox
		$UserToPermission = Read-Host "Inserisci il nome utente della persona o del gruppo al quale intendi dare accesso alla casella (esempio: mario.rossi@contoso.com o Mario Rossi)"

		#Get Permission Type
		$title = "TIPO PERMESSI"
		$message = "Che tipo di permesso vuoi applicare all'utente appena dichiarato?"
		$FullAccess = New-Object System.Management.Automation.Host.ChoiceDescription "&Full Access", ""
		$ReadOnly = New-Object System.Management.Automation.Host.ChoiceDescription "&Read Only", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($FullAccess, $ReadOnly)
		$PermissionType = $host.ui.PromptForChoice($title, $message, $options, 0)

		#Set the Permissions
		if ($PermissionType -eq 0) {
			try {
				""; ""; Write-Host "Inserisco i permessi Full Access sulla shared mailbox"; "";
				Add-MailboxPermission $SharedMailboxDisplayName -User $UserToPermission -AccessRights FullAccess -AutoMapping $true
			} catch {
				Write-Host "Non riesco ad applicare i permessi Full Access, riprovare." -f "Red"
				write-host $error[0]
				return ""
			}
		} else {
			try {
				""; "" Write-Host "Inserisco i permessi Read Only sulla shared mailbox"; "";
				Add-MailboxPermission $SharedMailboxDisplayName -User $UserToPermission -AccessRights ReadPermission -AutoMapping $true
			} catch {
				Write-Host "Non riesco ad applicare i permessi Read Only, riprovare." -f "Red"
				write-host $error[0]
				return ""
			}
		}

		try {
			Add-RecipientPermission $SharedMailboxDisplayName -Trustee $UserToPermission -AccessRights SendAs
		} catch {
			Write-Host "Non riesco ad applicare i permessi 'invia come', riprovare." -f "Red"
			write-host $error[0]
			return ""
		}

		#See if user wants to add additional permissions to mailbox
		$title = "ALTRI PERMESSI"
		$message = "Vuoi aggiungere ulteriori permessi alla casella creata?"
		$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Aggiungi permessi."
		$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
		$result = $host.ui.PromptForChoice($title, $message, $options, 0)
	}

	""; Write-Host "Imposto salvataggio copia email in uscita per $SharedMailboxUserName " -f "Yellow"
	Set-Mailbox $SharedMailboxUserName -MessageCopyForSentAsEnabled $True
	Set-Mailbox $SharedMailboxUserName -MessageCopyForSendOnBehalfEnabled $True

	""; Write-Host "All Done!" -f "Green"
	Write-Host "Riepilogo accessi alla casella di $SharedMailboxUserName " -f "Yellow"
	Get-MailboxPermission -Identity $SharedMailboxUserName | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,AccessRights | ft User,AccessRights | out-string
	Get-RecipientPermission $SharedMailboxUserName -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | ft Trustee, AccessRights | out-string
}

# Start script
. Main
