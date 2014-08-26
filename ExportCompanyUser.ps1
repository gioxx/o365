############################################################################################################################
# OFFICE 365: Export Company Users
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\ExportCompanyUsers.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		01-08-2014
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Export Company Users" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	$RicercaDominio = Read-Host "Dominio da analizzare (esempio: domain.tld) "
	
	try
	{
		""
		Write-Host "Ricerco le caselle con il dominio che mi hai richiesto, attendi." -foregroundcolor "yellow"
		$RicercaMailbox= Get-Mailbox -ResultSize unlimited | where {$_.EmailAddresses -like "*@" + $RicercaDominio}
		""
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