<#
	OFFICE 365: Get Message Trace (Single User or Mail Domain)
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.3
	Utilizzo:			.\MessageTrace.ps1
						(opzionale, passaggio dati da prompt) .\AddMailboxPermission.ps1 shared@contoso.com mario.rossi@contoso.com (oppure .\AddMailboxPermission.ps1 shared@contoso.com mario.rossi@contoso.com)
	Info:				http://gioxx.org/tag/o365-powershell
	Sources:			http://technet.microsoft.com/it-it/library/jj200704%28v=exchg.150%29.aspx
						http://community.office365.com/en-us/f/148/t/239828.aspx
						http://en.community.dell.com/techcenter/powergui/f/4834/t/19571829
						https://gallery.technet.microsoft.com/office/Office-365-Mail-Traffic-afa37da1
	Ultima modifica:	06-06-2019
	Modifiche:
		0.3- aggiungo delimitatore ";" all'export-CSV.
		0.2- rivisto lo script. Ora accetta parametri e permette di filtrare il singolo utente piuttosto che l'intero dominio. 
#>

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Domain, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $User
)

#Main
Function Main {

	""
	Write-Host "        Office 365: Get Message Trace" -f "Green"
	Write-Host "        ------------------------------------------"
	""
	
	if ( [string]::IsNullOrEmpty($Domain) -eq $false ) { 
		$TargetRicerca = "*@" + $Domain
		Write-Host "	Target di ricerca: *@$($Domain)" -f "Yellow" }
	if ( [string]::IsNullOrEmpty($User) -eq $false ) {
		$TargetRicerca = $User
		Write-Host "	Target di ricerca: $User" -f "Yellow" }
	if ( [string]::IsNullOrEmpty($Domain) -AND [string]::IsNullOrEmpty($User) ) { 
		$TargetRicerca = "*@contoso.com"
		Write-Host "	Dominio non specificato, imposto su contoso.com (DEMO)" -f "Yellow"; "";
		Write-Host "	Puoi terminare lo script con la combinazione CTRL + C" -f "Cyan" }
		
	""; Write-Host "	ATTENZIONE:" -f "red"
	Write-Host "	Specificare ora la data di inizio e di fine per scaricare i log," -f "Red"
	Write-Host "	dichiarandola come mm/gg/aaaa (americana), ad esempio 10/01/2014" -f "Red"
	Write-Host "	corrisponde al 10 gennaio 2014" -f "Red"; "";
	$DataInizio = Read-Host "Intervallo di analisi, data e ora di inizio (esempio: 10/01/2014 17:00)  "
	$DataFine = Read-Host "Intervallo di analisi, data e ora di fine   (esempio: 10/10/2014 19:00)  "
	
	$CurrentDate = Get-Date
	$CurrentDate = $CurrentDate.ToString('ddMMyyyy_hhmmss')
	if((Test-Path c:\temp) -eq 0) { new-item -type directory -path c:\temp }
	$NomeFile = $TargetRicerca -replace '[#?\{@*]', ''
	
	try
	{
		""; Write-Host "Ricerco tutte le mail inviate da $TargetRicerca, porta pazienza." -f "Yellow"
		$IngoingTotale = $null 
		$PagineIngoing = 1
		do 
			{ 
				Write-Host "Raccolta risultati - Pagina $PagineIngoing..." 
				$BloccoIngoingAttuale = Get-MessageTrace -SenderAddress $TargetRicerca -StartDate "$DataInizio" -EndDate "$DataFine" -PageSize 5000 -Page $PagineIngoing | Select Received,*Address,Subject
				<#	Nel caso non si voglia tracciare l'oggetto delle mail, togliere "Subject" nella Select, come di seguito:
					$BloccoIngoingAttuale = Get-MessageTrace -SenderAddress $TargetRicerca -StartDate "$DataInizio" -EndDate "$DataFine" -PageSize 5000 -Page $PagineIngoing | Select Received,*Address
				#>
				$PagineIngoing++
				$IngoingTotale += $BloccoIngoingAttuale
			} 
		until ($BloccoIngoingAttuale -eq $null)
		$IngoingTotale | Where-Object {$_} | Export-Csv C:\temp\$NomeFile-outgoing-$CurrentDate.csv -NoTypeInformation -Delimiter ";"
		
		""; Write-Host "Ricerco tutte le mail ricevute da $TargetRicerca, porta pazienza." -f "Yellow"
		$OutgoingTotale = $null 
		$PagineOutgoing = 1
		do 
			{ 
				Write-Host "Raccolta risultati - Pagina $PagineOutgoing..." 
				$BloccoOutgoingAttuale = Get-MessageTrace -RecipientAddress $TargetRicerca -StartDate "$DataInizio" -EndDate "$DataFine" -PageSize 5000 -Page $PagineOutgoing | Select Received,*Address,Subject
				$PagineOutgoing++
				$OutgoingTotale += $BloccoOutgoingAttuale
			} 
		until ($BloccoOutgoingAttuale -eq $null)
		$OutgoingTotale | Where-Object {$_} | Export-Csv C:\temp\$NomeFile-ingoing-$CurrentDate.csv -NoTypeInformation -Delimiter ";"
		
		""; Write-Host "Done." -f "green"
		Write-Host "Puoi trovare i file CSV rispettivamente in"
		Write-Host "  - C:\temp\$NomeFile-outgoing-$CurrentDate.csv"
		Write-Host "  - C:\temp\$NomeFile-ingoing-$CurrentDate.csv"
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

Function OutOfThere {}