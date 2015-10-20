############################################################################################################################
# OFFICE 365: Remove Mailbox Permission
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.5
# Utilizzo:				.\RemoveMailboxPermission.ps1
#						(opzionale, passaggio dati da prompt) .\RemoveMailboxPermission.ps1 -SourceMailbox shared@contoso.com -RemoveAccessTo mario.rossi@contoso.com
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		20-10-2015
# Modifiche:
#	0.5- Corretto if-else di richiesta informazioni da prompt.
#	0.4- Lo script accetta adesso i parametri passati da riga di comando (-SourceMailbox e -RemoveAccessTo)
#	0.3- Inclusa la rimozione dell'ACL GrantSendOnBehalfTo (prima non era prevista, aggiunta anche nello script di AddPermission)
#	0.2- Inclusa la rimozione dell'ACL SendAs (prima veniva effettuata la rimozione del solo FullAccess, rimaneva attivo il SendAs)
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $SourceMailbox, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $RemoveAccessTo 
)

""
Write-Host "        Office 365: Remove Mailbox Permission" -f "green"
Write-Host "        ------------------------------------------"

if (([string]::IsNullOrEmpty($SourceMailbox) -eq $true) -or ([string]::IsNullOrEmpty($RemoveAccessTo) -eq $true))
{
	#MANCANO I DETTAGLI DA PROMPT, LI RICHIEDO A VIDEO
	
	Write-Host "          ATTENZIONE:" -f "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -f "red"
	Write-Host "          nei dati richiesti qui di seguito" -f "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	""

    $SourceMailbox = Read-Host "Casella da modificare (esempio: sharedmailbox@domain.tld)            "
	$RemoveAccessTo = Read-Host "Utente al quale togliere l'accesso (esempio: mario.rossi@domain.tld) "
}

try
{
	""
	Write-Host "Rimuovo gli accessi di $RemoveAccessTo su $SourceMailbox ..." -f "yellow"
	Remove-MailboxPermission -Identity $SourceMailbox -User $RemoveAccessTo -AccessRights FullAccess -InheritanceType All
	Write-Host "Accesso rimosso per $RemoveAccessTo (su $SourceMailbox)" -f "green"
	""
	Write-Host "Tento rimozione parametro SendAs per $RemoveAccessTo (su $SourceMailbox)" -f "yellow"
	Remove-RecipientPermission $SourceMailbox -Trustee $RemoveAccessTo -AccessRights SendAs
	""
	Write-Host "Tento rimozione parametro GrantSendOnBehalfTo per $RemoveAccessTo (su $SourceMailbox)" -f "yellow"
	Set-Mailbox $SourceMailbox –Grantsendonbehalfto @{Remove="$RemoveAccessTo"}
	""
	Write-Host "All Done!" -f "green"
	Write-Host "Riepilogo accessi alla casella di $SourceMailbox " -f "yellow"
	Get-MailboxPermission -Identity $SourceMailbox
}
catch
{
	Write-Host "Non riesco ad elaborare l'operazione richiesta. Verificare i dettagli inseriti" -f "red"
	Write-Host $error[0]
	return ""
}