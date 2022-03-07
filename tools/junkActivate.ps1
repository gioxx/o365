<#	O365 PShell Snippet:	Activate Junk Email Protection
Autore (ver.-mod.):		    GSolone (0.1 ult.mod. 13/12/18)
Utilizzo:				          .\junkActivate.ps1 user@contoso.com
Info:					            https://gioxx.org/tag/o365-powershell
#>
Param( [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)][string] $SourceMailbox )
"";Write-Host "Attivo protezione antispam per $($SourceMailbox)" -f "yellow"
Set-MailboxJunkEmailConfiguration $SourceMailbox -Enabled $true
Get-MailboxJunkEmailConfiguration $SourceMailbox | Select Identity,Enabled
