############################################################################################################################
# OFFICE 365: Set multiple Ownership for all Security Groups
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetOwnershipSecurityGroups.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		28-11-2014
# Modifiche:			-
############################################################################################################################

#Main
Function Main {


# Modificare la variabile se necessario, includendo ulteriori indirizzi tra le virgolette e spaziati da una virgola
$GroupOwners = "admin01@domain.tld", "admin02@domain.tld"


	""
	Write-Host "        Office 365: Set multiple Ownership for all Security Groups" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          Lo script cerca tutti i gruppi di sicurezza" -foregroundcolor "Cyan"
	Write-Host "          attualmente presenti sul server che hanno nome" -foregroundcolor "Cyan"
	Write-Host "          'Security - *' e riapplica l'ownership secondo" -foregroundcolor "Cyan"
	Write-Host "          gli utenti specificati nella variabile all'interno" -foregroundcolor "Cyan"
	Write-Host "          dello script stesso." -foregroundcolor "Cyan"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	
	try
	{
		""
		Write-Host "Owners attualmente impostati nella variabile e che verranno riapplicati:" -foregroundcolor "yellow"
		$GroupOwners
		""
		Write-Host "Ricerco i gruppi di sicurezza presenti sul server Exchange, attendi." -foregroundcolor "yellow"
		$RicercaGruppi = Get-MsolGroup | where-object { $_.DisplayName -like "Security - *"}
		Write-Host "Done. Questi sono i gruppi di sicurezza attualmente presenti e trovati sul server:" -foregroundcolor "green"
		$RicercaGruppi | FT DisplayName,EmailAddress
		""
		Write-Host "Done. Applico l'ownership a tutti i gruppi, attendi." -foregroundcolor "green"
		$RicercaGruppi | ForEach-Object {Set-DistributionGroup $_.EmailAddress -ManagedBy $GroupOwners -BypassSecurityGroupManagerCheck}
		""
		Write-Host "Done." -foregroundcolor "green"
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