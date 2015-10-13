############################################################################################################################
# OFFICE 365: Create New Shared Mailbox
#----------------------------------------------------------------------------------------------------------------
# Original Author: 		Alan Byrne
# Version: 				1.0
# Last Modified Date: 	18/09/2012
# Last Modified By: 	Alan Byrne
#----------------------------------------------------------------------------------------------------------------
# Adattata da:			GSolone
# Versione:				0.7
# Utilizzo:				.\New-SharedMailbox.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		13-10-2015
# Modifiche:
#	0.7-ho commentato la forzatura del MicrosoftOnlineServicesID, non più necessaria
#	0.6-ho dovuto inserire un nuovo "sleep" prima di settare il MicrosoftOnlineServicesID perché in alcuni casi il server Exchange non trova immediatamente l'utente, cosa che invece succede già a 10 secondi di distanza dal comando di creazione casella.
#	0.5-correzioni minori
#	0.1/0-4-rimosso limite Shared a 5GB, forzata la richiesta dell'alias da utilizzare, rimossa richiesta permission Send As e Automap (entrambe dati per scontati e impostati di default), forzato il MicrosoftOnlineServicesID sull'indirizzo completo di posta elettronica
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: New Shared Mailbox" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$SharedMailboxUserName = Read-Host "Indirizzo (esempio: info@domain.tld)        "
	$SharedMailboxDisplayName = Read-Host "Nome casella (esempio: Contoso srl -Info)   "
	#Forzo richiesta alias
	$SharedMailboxAlias = Read-Host "Alias (esempio: Contososrl_info)            "
	
	#Create the new Shared mailbox
	try
	{
		""
		Write-Host "Creo la Shared Mailbox, attendi." -foregroundcolor "yellow"
		""
		New-Mailbox -Name $SharedMailboxDisplayName -Alias $SharedMailboxAlias -Shared -PrimarySMTPAddress $SharedMailboxUserName
		Start-Sleep -s 10
		#Forzo il MicrosoftOnlineServicesID (per evitare il default NOME@DOMINIO.onmicrosoft.com)
		#Set-Mailbox $SharedMailboxDisplayName -MicrosoftOnlineServicesID $SharedMailboxUserName
	}
	catch
	{
		Write-Host "Non riesco a creare la Shared Mailbox, riprovare." -foregroundcolor "red"
		write-host $error[0]
		return ""
	}
	
	Start-Sleep -s 10
	""
	Write-Host "Shared Mailbox creata correttamente: $SharedMailboxUserName" -foregroundcolor "green"
	""
	
	#Prompt user to add permissions to mailbox
	$title = "PERMESSI"
	$message = "Vuoi aggiungere adesso i permessi di accesso alla casella?"

	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Aggiungi permessi."
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

	$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
	
	while ($result -eq 0) 
	{
		""
		Write-Host ""
		#User to add to mailbox
		$UserToPermission = Read-Host "Inserisci il nome utente della persona o del gruppo al quale intendi dare accesso alla casella (esempio: mario.rossi@domain.tld o Mario Rossi)"
		
		#Get Permission Type
		$title = "TIPO PERMESSI"
		$message = "Che tipo di permesso vuoi applicare all'utente appena dichiarato?"

		$FullAccess = New-Object System.Management.Automation.Host.ChoiceDescription "&Full Access", ""
		$ReadOnly = New-Object System.Management.Automation.Host.ChoiceDescription "&Read Only", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($FullAccess, $ReadOnly)

		$PermissionType = $host.ui.PromptForChoice($title, $message, $options, 0)
		
		#Set the Permissions
		if ($PermissionType -eq 0)
		{
			try
			{
				""
				Write-Host ""
				Write-Host "Inserisco i permessi Full Access sulla shared mailbox"
				""
#		Blocco il controllo variabile Automap.
#				if ($AutoMap -eq 0)
#				{
					Add-MailboxPermission $SharedMailboxDisplayName -User $UserToPermission -AccessRights FullAccess -AutoMapping $true
#				}
#				else
#				{	
#					Add-MailboxPermission $SharedMailboxDisplayName -User $UserToPermission -AccessRights FullAccess -AutoMapping $false
#				}
			}
			catch
			{
				Write-Host "Non riesco ad applicare i permessi Full Access, riprovare." -foregroundcolor "red"
				write-host $error[0]
				return ""
			}
		}
		else
		{
			try
			{
				""
				Write-Host ""
				Write-Host "Inserisco i permessi Read Only sulla shared mailbox"
				""
#		Blocco il controllo variabile Automap.
#				if ($AutoMap -eq 0)
#				{
					Add-MailboxPermission $SharedMailboxDisplayName -User $UserToPermission -AccessRights ReadPermission -AutoMapping $true
#				}
#				else
#				{	
#					Add-MailboxPermission $SharedMailboxDisplayName -User $UserToPermission -AccessRights ReadPermission -AutoMapping $false
#				}
			}
			catch
			{
				Write-Host "Non riesco ad applicare i permessi Read Only, riprovare." -foregroundcolor "red"
				write-host $error[0]
				return ""
			}
		}
		
#		Blocco il controllo variabile Send As.
#		#Set sendas permissions if required
#		if ($SendAs -eq 0)
#		{
			try
			{
				Add-RecipientPermission $SharedMailboxDisplayName -Trustee $UserToPermission -AccessRights SendAs
			}
			catch
			{
				Write-Host "Non riesco ad applicare i permessi 'invia come', riprovare." -foregroundcolor "red"
				write-host $error[0]
				return ""
			}
#		}

		#See if user wants to add additional permissions to mailbox
		$title = "ALTRI PERMESSI"
		$message = "Vuoi aggiungere ulteriori permessi alla casella creata?"

		$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Aggiungi permessi."
		$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

		$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
	}
}

# Start script
. Main