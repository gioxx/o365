############################################################################################################################
# OFFICE 365: Remove Orphaned SID (Bulk, CSV)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.3
# Utilizzo:				.\RemoveOrphanedSID-CSV.ps1
#						(opzionale, modifica posizione file CSV) .\RemoveOrphanedSID-CSV.ps1 -csv C:\Export.CSV
#						(opzionale, analisi singola casella) .\RemoveOrphanedSID-CSV.ps1 -mbox shared@contoso.com
#						(opzionale, avvio rimozione) .\RemoveOrphanedSID-CSV.ps1 -action remove
#						I parametri da prompt possono essere concatenati.
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		18-12-2015
# Modifiche:			
#	0.3- prevedo concatenamento del file CSV da utilizzare con Action di Remove. Così facendo salto l'esportazione. Esempio di utilizzo: .\RemoveOrphanedSID-CSV.ps1 -csv C:\temp\OrphanedSID.csv -action remove. Corretti errori minori, modificata indentazione dello script. Richiedo ora cancellazione della lista CSV al termine di una Action di Remove.
#	0.2 rev1- correzioni minori (maggiori informazioni di utilizzo nel messaggio a video quando si lancia lo script).		
#	0.2- le condizioni di verifica fanno ora scomparire / comparire istruzioni a video. Prevedo l'utilizzo di un file CSV temporaneo nel caso in cui si vada a filtrare una sola mailbox con action remove. In caso contrario, in base a quanti e quali parametri utilizzo, mi comporto diversamente. Il file temporaneo (CSV) di esportazione SID orfani di singola casella, viene poi eliminato al termine dell'operazione.
#	0.1 rev1- ho solo corretto un errore in un messaggio di informazione.
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $CSV,
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Action,
	[Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Mbox
)

#Main
Function Main {

	################################################################################################
	# Puoi modificare il valore default di $ExportList per impostare un diverso nome del file CSV
	# che verrà esportato dallo script (solo ciò che c'è tra le virgolette, ad esempio
	# $ExportList = "C:\temp\export.csv" (per modificare anche la cartella di esportazione), 
	# oppure $ExportList = "Permessi.csv" per salvare il file nella stessa cartella dello script.
	
	if ([string]::IsNullOrEmpty($CSV))
		{ $ExportList = "C:\temp\OrphanedSID.csv" } else { $ExportList = $CSV }

	################################################################################################

	""
	Write-Host "        Office 365: Remove Orphaned SID (CSV)" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "         Lo script si occupa di analizzare, esportare e/o rimuovere i SID orfani"
	Write-Host "         rimasti memorizzati nelle Mailbox Permission sulle caselle presenti in Exchange"
	Write-Host "         (vedi commenti nello script per conoscere i parametri disponibili da prompt)"
	Write-Host "         Utilizzo: .\RemoveOrphanedSID-CSV.ps1 -csv C:\Export.CSV"
	Write-Host "                   .\RemoveOrphanedSID-CSV.ps1 -mbox shared@contoso.com"
	Write-Host "                   .\RemoveOrphanedSID-CSV.ps1 -action remove"
	Write-Host "         Vale concatenare i possibili parametri (es. -mbox shared@contoso.com -action remove)"
	Write-Host "                                                (es. -csv C:\temp\OrphanedSID.csv -action remove)"
	""
	
try
{
	
#######################################################################################################################
# E S T R A Z I O N E
#######################################################################################################################

	# Notifico posizione file CSV
	Write-Host "         File CSV in uso: $ExportList " -f "Green"; "";
	
	if ([string]::IsNullOrEmpty($Action)) {
		# Procedo con l'esportazione dei SID orfani di tutte le caselle di posta elettronica
		# (Nessuna Action specificata, quindi è certamente un'azione di esportazione).
		Write-Host "         L'esportazione potrebbe richiedere molti minuti"
		Write-Host "         (dipende dalla quantità di caselle di posta nel sistema)"
		""
	}
	
	Write-Host "         Premi un tasto per continuare (CTRL+C seguito da invio per terminare)."; "";
	[void][System.Console]::ReadKey($true)
	
	if ([string]::IsNullOrEmpty($Action)) {
	# Esporto i SID orfani SOLO se non è prevista una Action di Remove
		Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco ed esporto i SID in $ExportList , attendi..."
		
		if ([string]::IsNullOrEmpty($Mbox)) {
		# Nessuna mailbox specificata, estraggo i SID di tutte le caselle in Exchange
			Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {$_.accessrights -eq "FullAccess" -and $_.user -like "S-1-5-21*"} | Select-object identity,user | Export-CSV -NoTypeInformation $ExportList
		} else {
		# Casella di posta specificata, estraggo i SID presenti a bordo (se esistono)
			""
			Write-Host "         Analizzo la casella di posta $Mbox ..." -f "Yellow"
			Get-Mailbox $Mbox | Get-MailboxPermission | where {$_.accessrights -eq "FullAccess" -and $_.user -like "S-1-5-21*"} | Select-object identity,user
			""
			if ([string]::IsNullOrEmpty($Action)) {
				# Uso il file CSV di default o dichiarato nel caso in cui non ci sia un'action dichiarata
				Write-Host "         Esporto i risultati nel file CSV $ExportList ..." -f "Yellow"
				Get-Mailbox $Mbox | Get-MailboxPermission | where {$_.accessrights -eq "FullAccess" -and $_.user -like "S-1-5-21*"} | Select-object identity,user | Export-CSV -NoTypeInformation $ExportList
			} else {
				Write-Host "         Esporto i risultati nel file CSV temporaneo ..." -f "Yellow"
				# Altrimenti mi appoggio al file temporaneo che servirà a lanciare l'action in seguito
				Get-Mailbox $Mbox | Get-MailboxPermission | where {$_.accessrights -eq "FullAccess" -and $_.user -like "S-1-5-21*"} | Select-object identity,user | Export-CSV -NoTypeInformation C:\temp\OrphSidTemp.csv
			}
		}
		""; Write-Host "         Done." -f "Green"; "";
		# Chiedo se visualizzare i risultati esportati nel file CSV solo se NON specifico un'Action da prompt
		$message = "Devo aprire il file CSV $ExportList ?"
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$Excel = $host.ui.PromptForChoice("", $message, $options, 1)
		if ($Excel -eq 0) { Invoke-Item $ExportList }
		""
	}
	
#######################################################################################################################
# M O D I F I C A
#######################################################################################################################		

	# Remove dei SID (NON SPECIFICO PER SINGOLA MAILBOX), chiedo autorizzazione a procedere prima di continuare
	if ($Action -eq "remove" -and [string]::IsNullOrEmpty($Mbox)) {
		if ([string]::IsNullOrEmpty($CSV)) {
			# NESSUN CSV SPECIFICATO COME SORGENTE DATI
			# (prendo per buona l'ultimo file CSV specificato, ammesso che sia stato fatto, altrimenti il default)
			Write-Host "Prima di procedere con la rimozione ricordarsi di effettuare l'estrazione in CSV" -f "Red"; "";
			$message = "Posso procedere con la rimozione dei SID orfani esportati in $ExportList ?"
		} else {
			# CSV SPECIFICATO COME SORGENTE DATI
			# (quindi rimuovo i SID orfani delle caselle salvate nella lista specificata)
			Write-Host "File CSV da riga di comando: $ExportList" -f "Green"; "";
			$message = "Posso procedere con la rimozione dei SID orfani?"
		}
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$StartEngine = $host.ui.PromptForChoice("", $message, $options, 0)
		if ($StartEngine -eq 0) {
			import-csv $ExportList | ForEach-Object {
				Write-Progress -activity "Remove Orphaned SID" -status "Modifico $_.identity"
				Remove-MailboxPermission -Identity $_.identity -User $_.User -AccessRights FullAccess -InheritanceType All
			}
		}
		
		# Posso cancellare il file CSV utilizzato per rimuovere i SID orfani?
		""; ""; Write-Host "Il lavoro dovrebbe essere terminato." -f "Green"
		$message = "Posso procedere con la cancellazione del file $ExportList ?"
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$CancellazioneCSV = $host.ui.PromptForChoice("", $message, $options, 0)
		if ($CancellazioneCSV -eq 0) { Remove-Item $ExportList }
	}
	
	# Remove dei SID per singola Mailbox specificata a riga di comando
	if ($Action -eq "remove" -and [string]::IsNullOrEmpty($Mbox) -eq $false) {
		Write-Host "Pulisco i SID orfani di $Mbox ..." -f "Red"
		import-csv C:\temp\OrphSidTemp.csv | ForEach-Object {
			Write-Progress -activity "Remove Orphaned SID" -status "Modifico $_.identity"
			Remove-MailboxPermission -Identity $_.identity -User $_.User -AccessRights FullAccess -InheritanceType All
		}
		# Cancello il file CSV creato temporaneamente
		Remove-Item C:\temp\OrphSidTemp.csv
	}
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