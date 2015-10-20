############################################################################################################################
# OFFICE 365: List Mailboxes User Access
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.3
# Utilizzo:				.\ListMailboxesUserAccess-CSV.ps1
# Info:					http://gioxx.org/tag/o365-powershell
#						(opzionale, posizione CSV) .\ListMailboxesUserAccess-CSV.ps1 -CSV C:\Utenti.csv
#						(opzionale, filtro dominio) .\ListMailboxesUserAccess-CSV.ps1 -Domain contoso.com
#						(opzionale, numero caselle da analizzare) .\ListMailboxesUserAccess-CSV.ps1 -Count 10
# Ultima modifica:		19-10-2015
# Fonti utilizzate:		http://exchangeserverpro.com/list-users-access-exchange-mailboxes/
# Modifiche:		
#	0.3- introduco parametri da riga di comando per modificare posizione file CSV, filtrare un solo dominio di posta elettronica o limitare i risultati da ricercare.
#	0.2- corretti problemi di ricerca ACL sulle caselle degli utenti di posta elettronica. Ora lo script ricerca tutti i permessi Full-Access impostati sulle caselle di posta elettronica di Exchange (a prescindere che si tratti di Shared Mailbox o caselle di posta personali).
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $CSV,
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Domain,
    [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Count
)

#Main
Function Main {

	################################################################################################
	# Puoi modificare il valore $ExportList per impostare un diverso nome del file CSV che verrà
	# esportato dallo script (solo ciò che c'è tra le virgolette, ad esempio
	# $ExportList = "C:\temp\export.csv" (per modificare anche la cartella di esportazione), 
	# oppure $ExportList = "Permessi.csv" per salvare il file nella stessa cartella dello script.
	
	if ([string]::IsNullOrEmpty($CSV) -eq $true) { $ExportList = "C:\temp\MailboxPermissions.csv" } 
		else { $ExportList = $CSV }
		
	################################################################################################

	""
	Write-Host "        Office 365: List Mailboxes User Access" -foregroundcolor "Green"
	Write-Host "        ------------------------------------------"
	Write-Host "         Lo script elenca tutti i diritti Full Access per ogni casella di posta" -f "White"
	Write-Host "         presente su server Exchange, salvando i risultati su un file CSV" -f "White"
	Write-Host "         '" -f "White" -nonewline; Write-Host $ExportList -f "Green" -nonewline; Write-Host "'" -f "White"
	Write-Host "         (esegui lo script con parametro -CSV PERCORSOFILE.CSV per modificare)." -f "White"
	""
	
	<#do { $RicercaACL = Read-Host "Utente da cercare nei permessi (esempio: Mario Rossi) " } 
		while ($RicercaACL -eq [string]::empty)
	#>
	
	Write-Host "         Premi un tasto qualsiasi per continuare..."
	[void][System.Console]::ReadKey($true)
	
	try
	{
		""
		Write-Progress -Activity "Download dati da Exchange" -Status "Verifica permessi delle caselle di posta ed esportazione risultati in $ExportList, attendi..."
		
		<# Ricerca di test da lanciare manualmente per verifiche
		Write-Progress -Activity "DEBUG QUERY" -Status "Se stai leggendo questo messaggio c'è qualche problema, modifica lo script per inserire nuovamente la query di Debug nella parte commentata"
		Get-Mailbox -ResultSize 100 | Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation TestQuery.csv
		#>
		
		if ([string]::IsNullOrEmpty($Count) -eq $true) { $Count = "Unlimited" }		
		
		if ([string]::IsNullOrEmpty($Domain) -eq $false) {
			""
			Write-Host "         Dominio da analizzare: $Domain" -f "Yellow"
			Write-Host "         Counter analisi      : $Count" -f "Yellow"
			""
			# Se specificato .ps1 -Domain domain.tld, filtro solo il dominio richiesto (oltre le NT AUTHORITY\SELF e S-1-5*)
			Get-Mailbox -ResultSize $Count | where {$_.PrimarySmtpAddress -like "*@" + $Domain} | Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation $ExportList
		} else {
			""
			Write-Host "         Dominio da analizzare: All" -f "Yellow"
			Write-Host "         Counter analisi      : $Count" -f "Yellow"
			""
			# Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
			Get-Mailbox -ResultSize $Count | Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation $ExportList
		}
		
		Invoke-Item $ExportList
		Write-Host "Done." -f "Green"
		
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