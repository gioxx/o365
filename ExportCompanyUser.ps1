############################################################################################################################
# OFFICE 365: Export Company Users
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.4
# Utilizzo:				.\ExportCompanyUsers.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		26-10-2015
# Modifiche:			
#	0.4- Eliminato calcolo avanzamento download dati.
#	0.3- Correzione minore: la ricerca viene effettuata sullo specifico dominio in ingresso, un eventuale sottodominio deve essere dichiarato (esempio: contoso.com nella ricerca non mostrerà i risultati di dep1.contoso.com nell'eventualità dep1 fosse un suo sottodominio).
#	0.2- Stato di avanzamento in lettura / scrittura dati). Modificato il $_.EmailAddresses in $_.PrimarySmtpAddress per mettere la Company in base all'indirizzo di posta principale e non considerare eventuali alias
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Export Company Users" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	$RicercaDominio = Read-Host "Dominio da analizzare (esempio: domain.tld) "
	
	try
	{
		Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle con il dominio che mi hai richiesto, attendi..."
		
		#$RicercaMailbox= Get-Mailbox -ResultSize unlimited | where {$_.EmailAddresses -like "*@" + $RicercaDominio}
		$RicercaMailbox= Get-Mailbox -ResultSize Unlimited | where {$_.PrimarySmtpAddress -like "*@" + $RicercaDominio}
		
		$RicercaMailbox | FT DisplayName, UserPrincipalName
		Write-Host "Esporto l'elenco in C:\temp\$RicercaDominio.txt e apro il file (salvo errori)." -foregroundcolor "green"
		$ExportList = "C:\temp\$RicercaDominio.txt"
		
		$Today = [string]::Format( "{0:dd/MM/yyyy}", [datetime]::Now.Date )
		$TimeIs = (get-date).tostring('HH:mm:ss')		
		$RicercaMailbox | FT DisplayName, UserPrincipalName > $ExportList
		
		$a = Get-Content $ExportList
		$b = "Esportazione utenti $RicercaDominio - $Today alle ore $TimeIs"
#		Set-Content $ExportList –value $b, $a[0..18]
		Set-Content $ExportList –value $b, $a
		
		Invoke-Item $ExportList
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