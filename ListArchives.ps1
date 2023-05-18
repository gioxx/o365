<#
OFFICE 365: Get Online Archive List / Archive In-place (and stats)
-------------------------------------------------------------------------------------------------------------
Autore:							GSolone
Versione:						0.5
Utilizzo:						.\ListArchives.ps1
Info:								https://gioxx.org/tag/o365-powershell
Ultima modifica:		15-05-2023
Fonti utilizzate:		https://morgantechspace.com/2021/01/check-size-and-status-of-archive-mailbox-powershell.html
Modifiche:
  0.5- porto fuori dagli snippet e pulisco vecchio codice. Modifico per andare incontro a questo controllo suggerito qui da Microsoft: https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/archive-mailboxes/archivestatus-set-none (che ho già visto capitare sul tenant).
#>

$DataOggi = Get-Date -format yyyyMMdd
$CSV = "C:\temp\$($DataOggi)_ArchiveInPlace.csv"

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
        ArchiveStatus = $mbx.ArchiveStatus
        ArchiveName = $mbx.ArchiveName
        ArchiveState = $mbx.ArchiveState
        ArchiveMailboxSizeInGB = $size
        ArchiveWarningQuota =if ( $mbx.ArchiveDatabase -ne $null ) { $mbx.ArchiveWarningQuota } Else { $null }
        ArchiveQuota = if ( $mbx.ArchiveDatabase -ne $null ) { $mbx.ArchiveQuota } Else { $null }
        AutoExpandingArchiveEnabled = $mbx.AutoExpandingArchiveEnabled
        })
    } else {
      $size = 0
    }
  }
}
$Result | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Invoke-Item $CSV
