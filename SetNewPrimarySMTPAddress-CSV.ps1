############################################################################################################################
# OFFICE 365: Set New Primary SMTP Address (CSV)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\SetNewPrimarySMTPAddress-CSV.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		07-10-2015
# Modifiche:			
# 0.2-	lo script accetta ora il parametro CSV da riga di comando (esempio: .\SetNewPrimarySMTPAddress-CSV.ps1 C:\temp\Utenti.csv). Aggiunto il blocco di modifica MsolUserPrincipalName oltre l'indirizzo principale SMTP della casella di posta.
############################################################################################################################

#Verifica parametri da prompt
Param(
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]
    [string] $CSV
)

#Main
Function Main {

	""
	Write-Host "        Office 365: Add New PrimarySMTPAddress" -f "green"
	Write-Host "        ------------------------------------------"
	Write-Host "         Costruire il file CSV con in colonna 1 l'attuale indirizzo di posta elettronica" -f "white"
	Write-Host "         e in colonna 2 il nuovo indirizzo da far diventare primario." -f "white"
	Write-Host "         Il titolo della prima colonna dovrà essere " -nonewline; Write-Host "indirizzo_attuale" -f "yellow" -nonewline; Write-Host ", la seconda " -f "white" -nonewline; Write-Host "nuovo_indirizzo" -f "yellow"
	Write-Host "         Il file dovrà essere salvato come Utenti.csv e trovarsi nella stessa posizione di questo script" -f "white"
	""
	Write-Host "         In caso contrario sarà necessario lanciare lo script con parametro posizione del file CSV" -f "white"
	Write-Host "         esempio: .\SetNewPrimarySMTPAddress-CSV.ps1 C:\temp\Utenti.csv" -f "white"
	""
	Write-Host "         CSV di esempio:" -f "white"
	Write-Host "         indirizzo_attuale,nuovo_indirizzo" -f "gray"
	Write-Host "         test_1@contoso.onmicrosoft.com,test_1@contoso.com" -f "gray"
	Write-Host "         test_2@contoso.onmicrosoft.com,test_2@contoso.com" -f "gray"
	""
	Write-Host "		Premi un tasto qualsiasi per continuare..."
	[void][System.Console]::ReadKey($true)
	
	try
	{
		Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle elencate nel file CSV, attendi..."
		
		#NESSUN PARAMETRO PASSATO, CERCO UTENTI.CSV NELLA STESSA CARTELLA DELLO SCRIPT
		if ([string]::IsNullOrEmpty($CSV) -eq $true) { $CSV = ".\Utenti.csv" }
		
		#PROCEDO CON IL LAVORO DI MODIFICA INDIRIZZO PRIMARIO
			import-csv $CSV | ForEach-Object {
				$user = Get-Mailbox -Identity $_.indirizzo_attuale
				
				#DEBUG - INIZIO BLOCCO --------------------------------------------------------------------------------
					<#
					""; Write-Host "Se vedi questo testo è attivo il blocco debug, entra nello script e commentalo se necessario. Verifica user letto:" -f "red"
					$user
					#>
				#DEBUG - FINE BLOCCO ----------------------------------------------------------------------------------
				
				Write-Progress -activity "Modifica Primary SMTP Address" -status "Modifico $_.indirizzo_attuale"
				
				#RECUPERO DATI UTENTE
				$OldPrimarySMTPAddress = $_.indirizzo_attuale
				$NewPrimarySMTPAddress = $_.nuovo_indirizzo
				
				#MODIFICA IMPOSTAZIONI CASELLA DI POSTA (INDIRIZZO PRIMARIO)
				""; ""; Write-Host "		" -nonewline; Write-Host "Applicato nuovo indirizzo: $NewPrimarySMTPAddress" -b "green" -f "black"; ""
				$user.EmailAddresses += ("SMTP:$NewPrimarySMTPAddress")
				Set-Mailbox -Identity $user.Name -EmailAddresses $user.EmailAddresses
				
				#MODIFICA UTENTE IN EXCHANGE
				Set-MsolUserPrincipalName -UserPrincipalName $OldPrimarySMTPAddress -NewUserPrincipalName $NewPrimarySMTPAddress
			}
			
		""; ""; Write-Host "Script terminato, verifica da console che tutto sia andato liscio! :-)" -f "green"
		""
	}
	catch
	{
		""
		Write-Host "Errore nell'operazione, riprovare." -f "red"
		write-host $error[0]
		return ""
	}
	
}

# Start script
. Main