<#	O365 PShell Snippet:	Expand Mailbox SMTP Addresses
Autore (ver.-mod.):				GSolone (0.1 ult.mod. 15/7/16)
Utilizzo:									.\smtpExpand.ps1 user@contoso.com
Info:											https://gioxx.org/tag/o365-powershell
#>
Param( [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)][string] $SourceMailbox )
""; Write-Host "Indirizzi di posta elettronica associati a $SourceMailbox " -f "yellow"
Write-Host "(ricorda: 'SMTP' determina il Primary, 'smtp' è sempre un secondario)" -f "yellow"
# Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
Get-Recipient $SourceMailbox | Select Name -Expand EmailAddresses | where {$_ -like 'smtp*'}
""
