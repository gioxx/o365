############################################################################################################################
# OFFICE 365: Set User Company Field (Single User)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetCompanySingleUser.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		13-10-2015
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Set User Company Field (Single User)" -f "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -f "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -f "red"
	Write-Host "          nei dati richiesti qui di seguito" -f "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$RicercaUtente = Read-Host "Utente da modificare (esempio: mario.rossi@contoso.com)  "
	$RicercaCompany = Read-Host "Valore Company (esempio: Contoso S.r.l.)                "

	try
	{
		""
		Write-Host "Applico il valore Company a $RicercaUtente" -f "yellow"
		Set-User $RicercaUtente -Company $RicercaCompany
		""
		Write-Host "Modifica effettuata, verifica:" -f "green"
		Get-User $RicercaUtente | Select UserPrincipalName, Company
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
