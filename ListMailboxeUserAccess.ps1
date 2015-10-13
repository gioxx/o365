############################################################################################################################
# OFFICE 365: List Mailbox User Access
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\ListMailboxeUserAccess.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		22-09-2015
# Fonti utilizzate:		http://exchangeserverpro.com/list-users-access-exchange-mailboxes/
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: List Mailbox User Access" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "        Lo script elenca tutti i diritti Full Access di una qualsiasi" -f "white"
	Write-Host "        casella di posta specificata." -f "white"
	""
	
	do { $Mbox = Read-Host "Casella di posta da verificare (esempio: mailbox@contoso.com): " } 
		while ($Mbox -eq [string]::empty)
	
	try
	{
		""
		Write-Progress -Activity "Download dati da Exchange" -Status "Verifica permessi su $Mbox, attendi..."
		
		# Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
		Get-MailboxPermission $Mbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | ft User, 'Access Rights'
		
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