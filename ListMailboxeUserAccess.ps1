############################################################################################################################
# OFFICE 365: List Mailbox User Access
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\ListMailboxeUserAccess.ps1
# Info:					http://gioxx.org/tag/o365-powershell
#						(opzionale, passaggio dati da prompt) .\ListMailboxeUserAccess.ps1 shared@contoso.com
# Ultima modifica:		13-11-2015
# Fonti utilizzate:		http://exchangeserverpro.com/list-users-access-exchange-mailboxes/
# Modifiche:			
#	0.2- Verifico (e mostro a video) anche i permessi in SendAs sulla casella di posta specificata. Accetto parametri da riga di comando (così da saltare la richiesta della casella di posta elettronica da verificare).
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Mbox
)

#Main
Function Main {

	""
	Write-Host "        Office 365: List Mailbox User Access" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "        Lo script elenca tutti i diritti Full Access di una qualsiasi" -f "white"
	Write-Host "        casella di posta specificata." -f "white"
	""
	
if ( [string]::IsNullOrEmpty($Mbox) ) {
	#MANCANO I DETTAGLI DA PROMPT, LI RICHIEDO A VIDEO
	do { $Mbox = Read-Host "Casella di posta da verificare (esempio: mailbox@contoso.com): " } 
		while ($Mbox -eq [string]::empty)
	} else {
		""
		Write-Host "Casella di posta specificata a riga di comando: " -f "Yellow" -nonewline
		Write-Host "$Mbox" -f "Green"
	}
	
	try
	{
		""
		Write-Progress -Activity "Download dati da Exchange" -Status "Verifica permessi su $Mbox, attendi..."
		
		# Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
		Write-Host "Permessi in Full Access:" -f "Yellow"
		Get-MailboxPermission $Mbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | ft User, 'Access Rights'
		Write-Host "Permessi in Send As:" -f "Yellow"
		Get-RecipientPermission $Mbox -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | ft Trustee, AccessRights
		
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