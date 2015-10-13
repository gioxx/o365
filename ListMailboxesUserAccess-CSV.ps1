############################################################################################################################
# OFFICE 365: List Mailboxes User Access
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\ListMailboxesUserAccess-CSV.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		22-09-2015
# Fonti utilizzate:		http://exchangeserverpro.com/list-users-access-exchange-mailboxes/
# Modifiche:			
#	0.2- corretti problemi di ricerca ACL sulle caselle degli utenti di posta elettronica. Ora lo script ricerca tutti i permessi Full-Access impostati sulle caselle di posta elettronica di Exchange (a prescindere che si tratti di Shared Mailbox o caselle di posta personali).
############################################################################################################################

#Main
Function Main {

	################################################################################################
	# Puoi modificare questo valore per impostare un diverso nome del file CSV che verrà
	# esportato dallo script (solo ciò che c'è tra le virgolette, ad esempio
	# $ExportList = "C:\temp\export.csv" (per modificare anche la cartella di esportazione), 
	# oppure $ExportList = "Permessi.csv" per salvare il file nella stessa cartella dello script.
	$ExportList = "C:\temp\MailboxPermissions.csv"
	################################################################################################

	""
	Write-Host "        Office 365: List Mailboxes User Access" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "        Lo script elenca tutti i diritti Full Access per ogni casella di posta" -f "white"
	Write-Host "        presente su server Exchange, salvando i risultati su un file CSV" -f "white"
	Write-Host "        '" -f "white" -nonewline; Write-Host $ExportList -f "green" -nonewline; Write-Host "' nella stessa cartella di questo script." -f "white"
	""
	
	<#do { $RicercaACL = Read-Host "Utente da cercare nei permessi (esempio: Mario Rossi) " } 
		while ($RicercaACL -eq [string]::empty)
	#>
	
	Write-Host "        Premi un tasto qualsiasi per continuare..."
	[void][System.Console]::ReadKey($true)
	
	try
	{
		""
		Write-Progress -Activity "Download dati da Exchange" -Status "Verifica permessi delle caselle di posta ed esportazione risultati in $ExportList, attendi..."
		
		<# Ricerca di test da lanciare manualmente per verifiche
		Write-Progress -Activity "DEBUG QUERY" -Status "Se stai leggendo questo messaggio c'è qualche problema, modifica lo script per inserire nuovamente la query di Debug nella parte commentata"
		Get-Mailbox -ResultSize 100 | Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation TestQuery.csv
		#>
		
		# Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
		Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation $ExportList
		
		Invoke-Item $ExportList
		Write-Host "Done." -f "green"
		
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