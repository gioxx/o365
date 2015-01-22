############################################################################################################################
# OFFICE 365: Get Message Trace
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\MessageTrace.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		09-10-2014
# Modifiche:			-
############################################################################################################################
# Sources:				http://technet.microsoft.com/it-it/library/jj200704%28v=exchg.150%29.aspx
#						http://community.office365.com/en-us/f/148/t/239828.aspx
#						http://en.community.dell.com/techcenter/powergui/f/4834/t/19571829
#						https://gallery.technet.microsoft.com/office/Office-365-Mail-Traffic-afa37da1
#

#Main
Function Main {

	""
	Write-Host "        Office 365: Get Message Trace" -foregroundcolor "green"
	Write-Host "-------------------------------------------------------------------------------------------------"
	""
	$RicercaDominio = Read-Host "Dominio da analizzare (esempio: domain.tld) "
	""
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Specificare ora la data di inizio e di fine per scaricare i log," -foregroundcolor "red"
	Write-Host "          dichiarandola come mm/gg/aaaa (americana), ad esempio 10/01/2014" -foregroundcolor "red"
	Write-Host "          corrisponde al 10 gennaio 2014" -foregroundcolor "red"
	""
	$DataInizio = Read-Host "Intervallo di analisi, data e ora di inizio (esempio: 10/01/2014 17:00)  "
	$DataFine = Read-Host "Intervallo di analisi, data e ora di fine   (esempio: 10/10/2014 19:00)  "
	
	$TargetRicerca = "*@" + $RicercaDominio
	$CurrentDate = Get-Date
	$CurrentDate = $CurrentDate.ToString('ddMMyyyy_hhmmss')
	if((Test-Path c:\temp) -eq 0) { new-item -type directory -path c:\temp }
	
	try
	{
		""
		Write-Host "Ricerco tutte le mail inviate dal dominio $RicercaDominio, porta pazienza." -foregroundcolor "yellow"
		#Get-MessageTrace -SenderAddress $TargetRicerca -StartDate "$DataInizio" -EndDate "$DataFine" -PageSize 5000 | select Received,*Address | Export-Csv -Path "C:\temp\$RicercaDominio-outgoing-$CurrentDate.csv"
		$IngoingTotale = $null 
		$PagineIngoing = 1
		do 
			{ 
				Write-Host "Raccolta risultati - Pagina $PagineIngoing..." 
				$BloccoIngoingAttuale = Get-MessageTrace -SenderAddress $TargetRicerca -StartDate "$DataInizio" -EndDate "$DataFine" -PageSize 5000 -Page $PagineIngoing | Select Received,*Address
				$PagineIngoing++
				$IngoingTotale += $BloccoIngoingAttuale
			} 
		until ($BloccoIngoingAttuale -eq $null)
		$IngoingTotale | Where-Object {$_} | Export-Csv C:\temp\$RicercaDominio-outgoing-$CurrentDate.csv -NoTypeInformation
		
		""
		Write-Host "Ricerco tutte le mail ricevute dal dominio $RicercaDominio, porta pazienza." -foregroundcolor "yellow"
		#Get-MessageTrace -RecipientAddress $TargetRicerca -StartDate "$DataInizio" -EndDate "$DataFine" -PageSize 5000  | select Received,*Address | Export-Csv -Path "C:\temp\$RicercaDominio-ingoing-$CurrentDate.csv"
		$OutgoingTotale = $null 
		$PagineOutgoing = 1
		do 
			{ 
				Write-Host "Raccolta risultati - Pagina $PagineOutgoing..." 
				$BloccoOutgoingAttuale = Get-MessageTrace -RecipientAddress $TargetRicerca -StartDate "$DataInizio" -EndDate "$DataFine" -PageSize 5000 -Page $PagineOutgoing | Select Received,*Address
				$PagineOutgoing++
				$OutgoingTotale += $BloccoOutgoingAttuale
			} 
		until ($BloccoOutgoingAttuale -eq $null)
		$OutgoingTotale | Where-Object {$_} | Export-Csv C:\temp\$RicercaDominio-ingoing-$CurrentDate.csv -NoTypeInformation
		
		""
		Write-Host "Done." -foregroundcolor "green"
		Write-Host "Puoi trovare i file CSV rispettivamente in"
		Write-Host "  - C:\temp\$RicercaDominio-outgoing-$CurrentDate.csv"
		Write-Host "  - C:\temp\$RicercaDominio-ingoing-$CurrentDate.csv"
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