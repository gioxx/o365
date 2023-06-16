############################################################################################################################
# OFFICE 365: Add Distribution Group Member (CSV Bulk)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\AddDistributionGroupMember-CSV.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		09-12-2014
# Modifiche:			
#	0.2- modifiche minori
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Import Users from CSV to Distribution Group" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "          Prima di procedere, ricordarsi che in testa alla colonna del file CSV" -foregroundcolor "yellow"
	Write-Host "          dovrà essere indicato il testo 'Members', questo permetterà" -foregroundcolor "yellow"
	Write-Host "          allo script di inserire tutti gli utenti specificati (uno per riga) di" -foregroundcolor "yellow"
	Write-Host "          essere aggiunti al gruppo indicato precedentemente." -foregroundcolor "yellow"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$DistrGroup = Read-Host "Indirizzo o nome del gruppo (esempio: Messaggerie - Utenti)            "
	$CSVList = Read-Host "Posizione del file CSV (esempio: C:\Users\Utente\Desktop\elenco.csv)   "
	
	try
		{
			""
			Write-Host "Carico il contenuto del file CSV specificato." -foregroundcolor yellow
			Write-Host "L'importazione del file (salvo errori) potrebbe richiedere diverso tempo, attendi." -foregroundcolor yellow
			""
			Import-Csv $CSVList | foreach { Add-DistributionGroupMember -Identity "$DistrGroup" -Member $_.members }
			Write-Host "CSV importato nel gruppo $DistrGroup" -foregroundcolor green
			""
		}
	catch
		{
			Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
			write-host $error[0]
			return ""
		}
	
	""
	Write-Host "-------------------------------------------------------------------------------------------------" -foregroundcolor yellow
	""
	$title = ""
	$message = "Vuoi controllare chi fa ora parte del gruppo?"

	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Verifica adesso."
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

	$result = $host.ui.PromptForChoice($title, $message, $options, 0)
	if ($result -eq 0) { 
		""
		Write-Host "Questi sono gli utenti che ho trovato in $DistrGroup"
		Get-DistributionGroupMember $DistrGroup
	}
	
}

# Start script
. Main