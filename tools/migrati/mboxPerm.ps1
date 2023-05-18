<#	O365 PShell Snippet:	 Get Detailed Mailbox Permission
Autore (ver.-mod.):		     GSolone (0.4 ult.mod. 30/5/18)
Utilizzo:				           .\mboxPerm.ps1 user@contoso.com
Info:					             https://gioxx.org/tag/o365-powershell
#>
Param( [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)][string] $SourceMailbox )
""; Write-Host "Riepilogo accessi alla casella di $SourceMailbox " -f "yellow"
# Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
Get-MailboxPermission -Identity $SourceMailbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,AccessRights
Write-Host "--------------------------------------------------------" -f "Yellow"
Get-RecipientPermission $SourceMailbox -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | Select Identity, Trustee, AccessRights | out-string
Write-Host "GrantSendOnBehalfTo per $($SourceMailbox)" -f "Yellow"
Get-Mailbox $SourceMailbox | Select -Expand GrantSendOnBehalfTo
