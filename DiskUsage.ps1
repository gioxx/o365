<#
 OFFICE 365: Check Mailbox Size on Exchange (Disk Usage)
 ----------------------------------------------------------------------------------------------------------------
	Autore:					GSolone
	Versione:				0.1 rev1
	Utilizzo:				.\DiskUsage.ps1
							opzionale, modifica posizione export CSV: .\DiskUsage.ps1 C:\temp\Clutter.csv
							opzionale, singola mailbox specificata: .\DiskUsage.ps1 -Mailbox mario.rossi@contoso.com
	Info:					http://gioxx.org/tag/o365-powershell
	Fonti utilizzate:		http://www.morgantechspace.com/2015/09/find-office-365-mailbox-size-with-powershell.html
							https://nchrissos.wordpress.com/2013/06/17/reporting-mailbox-sizes-on-microsoft-exchange-2010/
	DEBUG per singolo nodo: 
							Get-Mailbox -ResultSize 20 | Get-MailboxStatistics | where { $_.OriginatingServer -match "eurprd07.prod.outlook.com"} | 
	Bug conosciuti:			se lo script trova una omonimia nel sistema, restituisce errore di tipo: 'La cassetta postale specificata "mario.rossi" non � univoca.'
	Ultima modifica:		04-10-2016
	Modifiche:				
		0.1 rev1-	correzioni minori.
#>

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $CSV,
	[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Mailbox
)

<#
	Puoi modificare il valore $CSV per impostare un diverso nome del file CSV che verr�
	esportato dallo script (solo ci� che c'� tra le virgolette, ad esempio
	$CSV = "C:\temp\CSV.csv" (per modificare anche la cartella di esportazione), 
	oppure $CSV = "Permessi.csv" per salvare il file nella stessa cartella dello script.
	ATTENZIONE: utilizza (per comodit�) nomi diversi nel caso in cui lo script esporti i permessi
				delle caselle ShaRed piuttosto che quelle personali.
#>
	$DataOggi = Get-Date -format yyyyMMdd
	if ([string]::IsNullOrEmpty($CSV) -eq $true) {
		$CSV = "C:\temp\DiskUsage_$DataOggi.csv"
	}

""
Write-Host "        Office 365: Check Mailbox Size on Exchange (Disk Usage)" -f "green"
Write-Host "        ------------------------------------------"

# Mailbox non specificata, estrazione dati completa
if ([string]::IsNullOrEmpty($Mailbox) -eq $true) {
	Write-Host "          ATTENZIONE:" -f "Red"
	Write-Host "          L'operazione pu� richiedere MOLTO tempo, dipende dal numero di utenti"
	Write-Host "          da verificare e modificare all'interno della Directory, porta pazienza!"
	""
	Write-Host "          Per modificare la posizione del file CSV esportato, rilancia lo script con parametro"
	Write-Host "          -CSV C:\export.csv" -nonewline -f "Yellow"; Write-Host " (es. " -nonewline; Write-Host ".\DiskUsage.ps1 -CSV C:\export.csv" -nonewline -f "Yellow"; Write-Host ")"
	""
	Write-Host "          Per analizzare una singola mailbox, rilancia lo script con parametro"
	Write-Host "          -Mailbox mario.rossi@contoso.com" -nonewline -f "Yellow"; Write-Host " (es. " -nonewline; Write-Host ".\DiskUsage.ps1 -Mailbox mario.rossi@contoso.com" -nonewline -f "Yellow"; Write-Host ")"
	Write-Host "        ------------------------------------------"
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
		""; Write-Host "		File CSV di destinazione: " -nonewline; Write-Host $CSV -f "Yellow"; ""
		Write-Host "		A long time left, grab a Snickers!" -f "Yellow"
		Write-Progress -Activity "Download dati da Exchange" -Status "Scarico occupazione disco delle caselle registrate nel sistema..."
		# Analisi occupazione Mailbox, sort, salvataggio su file
		Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics |
		Select-Object -Property @{label="User";expression={$_.DisplayName}},
		@{label="Total Messages";expression= {$_.ItemCount}},
		@{label="Total Size (GB)";expression={[math]::Round(`
			# Trasformo in GB
			($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}} |
		Sort "Total Size (GB)" -Descending > $CSV
		
		""
		Write-Host "Ho terminato l'esportazione dei dati." -f "Green"
		
		# Chiedo se visualizzare il file CSV Generato
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Apri ora il file CSV"
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Termina lo script senza aprire il file CSV"
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$OpenCSV = $host.ui.PromptForChoice("Devo aprire il file CSV generato?", $message, $options, 1)
		if ($OpenCSV -eq 0) { Invoke-Item $CSV }
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -f "Red"
		write-host $error[0]
		return ""
	}
	
} else {
	""
	Write-Host "        Mailbox da analizzare specificata: " -nonewline; Write-Host $Mailbox -f "Yellow"
	Get-Mailbox $Mailbox | Get-MailboxStatistics |
		Select-Object -Property @{label="User";expression={$_.DisplayName}},
		@{label="Total Messages";expression= {$_.ItemCount}},
		@{label="Total Size (GB)";expression={[math]::Round(`
			# Trasformo in GB
			($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}}
}