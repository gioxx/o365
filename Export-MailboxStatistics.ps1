<#
OFFICE 365: Export Mailbox Statistics
----------------------------------------------------------------------------------------------------------------
Autore originale:		Sconosciuto
Autore:					    GSolone
Versione:				    0.1
Utilizzo:				    .\Export-MailboxStatistics.ps1
Info:					      https://gioxx.org/tag/o365-powershell
Ultima modifica:		22-02-2022
Modifiche:
#>

$DataOggi = Get-Date -format yyyyMMdd
$CSV = "C:\temp\$($DataOggi)_MailboxSize.csv"

""; Write-Host "Esporto i dati, attendi" -f "Yellow"; "";
Get-Mailbox -ResultSize Unlimited | where {$_.RecipientType -eq "UserMailbox"} | Select-Object DisplayName,servername,database,RecipientTypeDetails,PrimarySmtpAddress,@{Name='TotalItemSize(GB)'; expression={[math]::Round((((Get-MailboxStatistics $_.PrimarySmtpAddress).TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},@{Name='ItemCount'; expression={(Get-MailboxStatistics $_.PrimarySmtpAddress).ItemCount}},IssueWarningQuota, ProhibitSendQuota | Export-Csv $CSV -Append -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Invoke-Item $CSV
