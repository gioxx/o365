<#	O365 PShell Snippet:	Get Online Archive List
Autore (ver.-mod.):		    GSolone (0.3 ult.mod. 4/2/22)
Utilizzo:				          .\ListArchives.ps1
Info:					            https://gioxx.org/tag/o365-powershell
#>
""; Write-Host "Archivi presenti su Exchange:"; "";
Write-Host "Utenti" -f "Yellow"
Get-Mailbox -ResultSize unlimited | Where {$_.archivedatabase -ne $null -AND $_.RecipientTypeDetails -eq 'UserMailbox'} | Select Name,Alias,PrimarySmtpAddress,Database,ProhibitSendQuota,ExternalDirectoryObjectId | Out-GridView
""; Write-Host "Shared Mailbox" -f "Yellow"; ""
Get-Mailbox -ResultSize unlimited | Where {$_.archivedatabase -ne $null -AND $_.RecipientTypeDetails -eq 'SharedMailbox'} | Select Name,Alias,PrimarySmtpAddress,Database,ProhibitSendQuota,ExternalDirectoryObjectId | Out-GridView
