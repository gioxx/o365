############################################################################################################################
# OFFICE 365: Export Shared Mailbox Permission (List Users with Full Access or Send As)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\ListSharedUsers.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		08-08-2014
# Modifiche:			-
############################################################################################################################
# ATTENZIONE:
# Lo script è ancora "sporco" e può generare errori di risoluzione nomi in console Powershell.
# Fa il suo lavoro, talvolta non riesce a completarlo ma i dati estratti sono corretti, prima o poi cercherò di capire se è possibile filtrare utenti e gruppi non "risolvibili" (prima o poi eh, senza fretta!)
############################################################################################################################

""
Write-Host "        Office 365: Export Shared Mailbox Permission" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          IN SVILUPPO:" -foregroundcolor "red"
Write-Host "          Lo script è ancora 'sporco' e può generare errori di" -foregroundcolor "red"
Write-Host "          risoluzione nomi in console Powershell. I risultati restituiti" -foregroundcolor "red"
Write-Host "          sono in ogni caso validi." -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"
$RicercaDominio = Read-Host "Dominio da analizzare (esempio: domain.tld) "
$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | where {$_.EmailAddresses -like "*" + $RicercaDominio} | Select Identity, Alias, DisplayName, DistinguishedName, PrimarySmtpAddress

ForEach ($Mailbox in $Mailboxes) {
	#$SendAs = Get-RecipientPermission $Mailbox.DistinguishedName | ? {$_.AccessRights -eq "SendAs" -and $_.trustee -ne "NT AUTHORITY\SELF" -and !$_.IsInherited} | % {$_.trustee} | get-user | % {$_.name}
	#$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User} | get-user | % {$_.name}
	
	$SendAs = Get-RecipientPermission $Mailbox.DistinguishedName | ? {$_.AccessRights -eq "SendAs" -and $_.trustee -ne "NT AUTHORITY\SELF" -and !$_.IsInherited} | % {$_.trustee} | get-user | % {$_.UserPrincipalName}	
	$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User} | get-user | % {$_.UserPrincipalName}

	if ([string]$sendas -ne [string]$fullaccess)
		{ 
			""
			Write-Host $Mailbox.DisplayName -f green -nonewline; Write-Host " (" -f green -nonewline; Write-Host $Mailbox.PrimarySmtpAddress -f green -nonewline; Write-Host ")" -f green;
			Write-Host "Send As    : " -f yellow -nonewline; Write-Host $sendas
			Write-Host "Full Access: " -f yellow -nonewline; Write-Host $fullaccess
		}  else { 
			""
			Write-Host $Mailbox.DisplayName -f green -nonewline; Write-Host " (" -f green -nonewline; Write-Host $Mailbox.PrimarySmtpAddress -f green -nonewline; Write-Host ")" -f green;
			Write-Host "Send As    : " -f yellow -nonewline; Write-Host $sendas
			Write-Host "Full Access: " -f yellow -nonewline; Write-Host $fullaccess
		}
}

""