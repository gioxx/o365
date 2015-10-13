############################################################################################################################
# OFFICE 365: Bulk Set City Field (Based on Primary SMTP Address)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetCityBulk.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		09-06-2014
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Bulk Set City Field (Based on Primary SMTP Address)" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$RicercaDominio = Read-Host "Dominio da analizzare (esempio: domain.tld) "
	$RicercaCity = Read-Host "Valore City (esempio: Assago)               "
	
	try
	{
		""
		Write-Host "Ricerco le caselle con il dominio che mi hai richiesto, attendi." -foregroundcolor "yellow"
		$RicercaMailbox= Get-Mailbox -ResultSize Unlimited | where {$_.PrimarySmtpAddress -like "*" + $RicercaDominio}
		Write-Host "Applico il valore City alle utenze rilevate ..." -foregroundcolor "yellow"
		""
		$RicercaMailbox | ForEach {Write-Host "Utente: $_"; Set-User $_.UserPrincipalName -City $RicercaCity}
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