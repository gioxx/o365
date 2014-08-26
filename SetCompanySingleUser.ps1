############################################################################################################################
# OFFICE 365: Set User Company Field (Single User)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetCompanySingleUser.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		29-05-2014
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Set User Company Field (Single User)" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$RicercaUtente = Read-Host "Utente da modificare (esempio: mario.rossi@domain.tld)  "
	$RicercaCompany = Read-Host "Valore Company (esempio: Emmelibri S.r.l.)              "
	
	try
	{
		""
		Write-Host "Applico il valore Company a $RicercaUtente" -foregroundcolor "yellow"
		Set-User $RicercaUtente -Company $RicercaCompany
		""
		Write-Host "Modifica effettuata, verifica:" -foregroundcolor "green"
		Get-User $RicercaUtente | Select UserPrincipalName, Company
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