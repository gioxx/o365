############################################################################################################################
# OFFICE 365: Remove Distribution Group Member (recursive)
#----------------------------------------------------------------------------------------------------------------
# Autore:							GSolone
# Versione:						0.1
# Utilizzo:						.\RemoveDistributionGroupMember.ps1
# Info:								https://gioxx.org/tag/o365-powershell
# Ultima modifica:		21-04-2015
# Modifiche:
############################################################################################################################

#Main
Function Main {
	""; Write-Host "        Office 365: Remove Distribution Group Member" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$DistrGroup = Read-Host "Indirizzo o nome del gruppo (esempio: Messaggerie - Utenti)            "
	$result = 0
	while ($result -eq 0)
	{
		try {
			$UsrGroup = Read-Host "Utente da rimuovere (esempio: mario.rossi/mario rossi/user@contoso.com) "
			Remove-DistributionGroupMember -Identity $DistrGroup -Member $UsrGroup
			Write-Host "Done." -f "Green"; "";

			$title = ""
			$message = "Vuoi rimuovere altri utenti dal gruppo?"
			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Rimuovi utente."
			$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
			$result = $host.ui.PromptForChoice($title, $message, $options, 0)
			""
		} catch {
			Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
			write-host $error[0]
			return ""
		}
	}

	""; Write-Host "-------------------------------------------------------------------------------------------------" -f "Yellow"; "";
	$title = ""
	$message = "Vuoi controllare chi fa ora parte del gruppo?"

	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Verifica adesso."
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

	$result = $host.ui.PromptForChoice($title, $message, $options, 0)
	if ($result -eq 0) {
		""; Write-Host "Questi sono gli utenti che ho trovato in $DistrGroup"
		Get-DistributionGroupMember $DistrGroup
	}

}

# Start script
. Main
