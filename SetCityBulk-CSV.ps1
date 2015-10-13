############################################################################################################################
# OFFICE 365: Set User City Field (Bulk, CSV)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetCitySingleBulk-CSV.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		15-05-2015
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Set User City Field (Bulk, CSV)" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "        Costruire il file CSV con in colonna 1 l'attuale indirizzo di posta elettronica" -f "white"
	Write-Host "        e in colonna 2 il relativo campo City." -f "white"
	Write-Host "        Il titolo della prima colonna dovrà essere " -nonewline; Write-Host "mailbox_utente" -f "yellow" -nonewline; Write-Host ", la seconda " -f "white" -nonewline; Write-Host "campo_city" -f "yellow"
	Write-Host "        Il file dovrà essere salvato come City.csv e trovarsi nella stessa posizione di questo script" -f "white"
	""
	Write-Host "        CSV di esempio (puoi fare copia e incolla da qui):" -f "white"
	Write-Host "        mailbox_utente,campo_city" -f "gray"
	Write-Host "        test_1@contoso.onmicrosoft.com,Buccinasco" -f "gray"
	Write-Host "        test_2@contoso.onmicrosoft.com,Assago" -f "gray"
	""
	Write-Host "		Premi un tasto qualsiasi per continuare..."
	[void][System.Console]::ReadKey($true)
	
	try
	{
		Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle elencate nel file CSV, attendi..."
		import-csv .\City.csv | ForEach-Object {
			
			Write-Progress -activity "Modifica campo City utente" -status "Modifico $_.mailbox_utente"
			Set-user $_.mailbox_utente -City $_.campo_city
			
			}
			
		""; ""; Write-Host "Script terminato, verifica da console che tutto sia andato liscio! :-)" -f "green"
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