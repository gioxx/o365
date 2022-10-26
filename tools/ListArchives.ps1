<#	O365 PShell Snippet:	Get Online Archive List
Autore (ver.-mod.):		    GSolone (0.4 ult.mod. 28/9/22)
Utilizzo:				          .\ListArchives.ps1
Info:					            https://gioxx.org/tag/o365-powershell
Credits:                  https://morgantechspace.com/2021/01/check-size-and-status-of-archive-mailbox-powershell.html
#>

$DataOggi = Get-Date -format yyyyMMdd
$CSV = "C:\temp\$($DataOggi)_ArchiveInPlace.csv"

<#
""; Write-Host "Esporto i dati, attendi" -f "Yellow"; "";
$ArchiveMbox = Get-Mailbox -ResultSize Unlimited | Where {$_.ArchiveDatabase -ne $null}
$ArchiveMbox | Foreach {
 Get-Mailbox -Archive $_.UserPrincipalName | Select-Object DisplayName,
 servername,database,
 RecipientTypeDetails,PrimarySmtpAddress,ArchiveStatus,ArchiveName,ArchiveState
 @{Name='TotalItemSize(GB)'; expression={[math]::Round((((Get-MailboxStatistics $_.UserPrincipalName).TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},
 IssueWarningQuota,ProhibitSendQuota,AutoExpandingArchiveEnabled |
 Export-Csv $CSV -Append -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}
Invoke-Item $CSV
#>

$Result=@()
$mailboxes = Get-Mailbox -ResultSize Unlimited
$totalmbx = $mailboxes.Count
$ProcessedCount=0
$mailboxes | ForEach-Object {
  $ProcessedCount++
  $mbx = $_
  $size = $null

  Write-Progress -Activity "Analisi in corso:" -Status "$ProcessedCount utenti di $totalmbx" -PercentComplete (($ProcessedCount / $totalmbx) * 100)

  if ($mbx.ArchiveDatabase -ne $null) {
    $mbs = Get-MailboxStatistics $mbx.UserPrincipalName -Archive
    if ($mbs.TotalItemSize -ne $null) {
      $size = [math]::Round(($mbs.TotalItemSize.Value.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)
      $Result += New-Object -TypeName PSObject -Property $([ordered]@{
        UserName = $mbx.DisplayName
        UserPrincipalName = $mbx.UserPrincipalName
        ArchiveStatus =$mbx.ArchiveStatus
        ArchiveName =$mbx.ArchiveName
        ArchiveState =$mbx.ArchiveState
        ArchiveMailboxSizeInGB = $size
        ArchiveWarningQuota=if ($mbx.ArchiveStatus -eq "Active") {$mbx.ArchiveWarningQuota} Else { $null}
        ArchiveQuota = if ($mbx.ArchiveStatus -eq "Active") {$mbx.ArchiveQuota} Else { $null}
        AutoExpandingArchiveEnabled=$mbx.AutoExpandingArchiveEnabled
        })
    } else {
      $size = 0
    }
  }
}
$Result | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Invoke-Item $CSV
