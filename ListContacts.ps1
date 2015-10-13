############################################################################################################################
# OFFICE 365: Search and show contacts & aliases (starting from a mail domain)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\ListContactsAlias.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		11-09-2015
# Modifiche:			
#	0.2- ho incluso nella ricerca anche gli alias delle caselle di posta di Exchange, così da riuscire ad individuare possibili altri indirizzi presenti in Exchange ma non dichiarati come contatti (valgono sia gli alias, sia i Primary SMTP Address).
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Search and show contacts & aliases" -f "green"
	Write-Host "        --------------------------------------------------"
	Write-Host "          ATTENZIONE:" -f "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -f "red"
	Write-Host "          nei dati richiesti qui di seguito" -f "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$RicercaDominio = Read-Host "Dominio da analizzare (esempio: domain.tld) "
	
	try
	{
		Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco i contatti con il dominio richiesto..."
		$RicercaContatti= Get-Contact -ResultSize Unlimited | where {$_.WindowsEmailAddress -like "*@" + $RicercaDominio}
		""
		Write-Host "Risultato della ricerca:" -f green
		$RicercaContatti | ft DisplayName, WindowsEmailAddress
		""
		
		# Ricerca degli alias (solo su richiesta)
		$title = ""
		$message = "Vuoi che estenda la ricerca anche negli alias di posta? (default: NO)"

		$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Verifica adesso."
		$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

		$result = $host.ui.PromptForChoice($title, $message, $options, 1)
		if ($result -eq 0) { 
			""
			Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco gli alias di posta con il dominio richiesto..."
			$RicercaMailbox= Get-Mailbox -ResultSize Unlimited | where {$_.EmailAddresses -like "*@" + $RicercaDominio}
			$RicercaMailbox | ft DisplayName, WindowsEmailAddress
		}
		""
		
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -f "red"
		write-host $error[0]
		return ""
	}
	
}

# Start script
. Main