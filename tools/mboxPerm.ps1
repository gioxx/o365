#####################################################################
# O365 PShell Snippet:	Get Detailed Mailbox Permission				#
# Autore (ver.-mod.):	GSolone (0.1 ult.mod. 15/1/16)				#
# Utilizzo:				.\mboxPerm.ps1 user@contoso.com				#
# Info:					http://gioxx.org/tag/o365-powershell		#
#####################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $SourceMailbox
)

""
Write-Host "Riepilogo accessi alla casella di $SourceMailbox " -f "yellow"
	# Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
	Get-MailboxPermission -Identity $SourceMailbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | ft User, "Access Rights" | out-string
	Get-RecipientPermission $SourceMailbox -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | ft Trustee, AccessRights | out-string