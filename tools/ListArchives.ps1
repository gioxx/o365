<#	O365 PShell Snippet:	Get Online Archive List
Autore (ver.-mod.):		    GSolone (0.4 ult.mod. 28/9/22)
Utilizzo:				          .\ListArchives.ps1
Info:					            https://gioxx.org/tag/o365-powershell
#>

$DataOggi = Get-Date -format yyyyMMdd
$CSV = "C:\temp\$($DataOggi)_ArchiveInPlace.csv"

""; Write-Host "Esporto i dati, attendi" -f "Yellow"; "";
Get-Mailbox -ResultSize Unlimited | Where {$_.ArchiveDatabase -ne $null} |
Select-Object DisplayName,
servername,database,
RecipientTypeDetails,PrimarySmtpAddress,
@{Name='TotalItemSize(GB)'; expression={[math]::Round((((Get-MailboxStatistics $_.PrimarySmtpAddress).TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},
@{Name='ItemCount'; expression={(Get-MailboxStatistics $_.PrimarySmtpAddress).ItemCount}},
IssueWarningQuota, ProhibitSendQuota |
Export-Csv $CSV -Append -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Invoke-Item $CSV
