############################################################################################################################
# OFFICE 365: Set User City Field (Bulk, CSV)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1 rev1
# Utilizzo:				.\SetCitySingleBulk-CSV.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		10-03-2016
# Modifiche:
#	0.1 rev1: aggiunta funzione di Pausa per evitare di intercettare il tasto CTRL.
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
	
	Function Pause($M="Premi un tasto continuare (CTRL+C per annullare)") {
		If($psISE) {
			$S=New-Object -ComObject "WScript.Shell"
			$B=$S.Popup("Fai clic su OK per continuare.",0,"In attesa dell'amministratore",0)
			Return
		}
		Write-Host -NoNewline $M;
		$I=16,17,18,20,91,92,93,144,145,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183;
		While($K.VirtualKeyCode -Eq $Null -Or $I -Contains $K.VirtualKeyCode) {
			$K=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}
	}
	Pause
	
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