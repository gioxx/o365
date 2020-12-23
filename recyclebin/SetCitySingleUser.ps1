############################################################################################################################
# OFFICE 365: Set User City Field (Single User)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetCitySingleUser.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		22-01-2015
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Set User City Field (Single User)" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$RicercaUtente = Read-Host "Utente da modificare (esempio: mario.rossi@domain.tld)  "
	$RicercaCompany = Read-Host "Valore City          (esempio: Milano)                  "
	
	try
	{
		""
		Write-Host "Applico il valore City a $RicercaUtente" -foregroundcolor "yellow"
		Set-User $RicercaUtente -City $RicercaCompany
		""
		Write-Host "Modifica effettuata, verifica:" -foregroundcolor "green"
		Get-User $RicercaUtente | Select UserPrincipalName, City
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