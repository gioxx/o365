<#
	O365 PShell Snippet: Get Online Archive List
	----------------------------------------------------------------------------------------------------------------
	Autore:		GSolone (0.2 ult.mod. 5/2/20)
	Utilizzo:	.\ListArchives.ps1
	Info:		http://gioxx.org/tag/o365-powershell
#>

""; Write-Host "Archivi presenti su Exchange:"; "";

<#
Write-Host "Utenti" -f "Yellow"
Get-Mailbox -ResultSize unlimited -Filter { ArchiveStatus -Eq "Active" -AND RecipientTypeDetails -eq 'UserMailbox'}
""; Write-Host "Shared Mailbox" -f "Yellow"; ""
Get-Mailbox -ResultSize unlimited -Filter { ArchiveStatus -Eq "Active" -AND RecipientTypeDetails -eq 'SharedMailbox'}
#>

Write-Host "Utenti" -f "Yellow"
Get-Mailbox -ResultSize unlimited | Where {$_.archivedatabase -ne $null -AND $_.RecipientTypeDetails -eq 'UserMailbox'}
""; Write-Host "Shared Mailbox" -f "Yellow"; ""
Get-Mailbox -ResultSize unlimited | Where {$_.archivedatabase -ne $null -AND $_.RecipientTypeDetails -eq 'SharedMailbox'}