<#
.SYNOPSIS
    Script obtains mailbox statistics from an Exchange tenant on M365
   
.DESCRIPTION
    Script obtains disk space occupancy statistics of mailboxes in an Exchange tenant (Microsoft 365) and collects them in a CSV file.
    The exported data are the Display Name, Server Name, Database, Recipient Type (personal mailbox or shared mailbox), Primary SMTP Address, Issue Warning Quota and Prohibit Send Quota.
    Script calculates disk space occupancy in GB.

.NOTES
    Filename: Export-MailboxStatistics.ps1
    Version: 0.4
    Last modified: 15-05-2023
    Author: GSolone
    Blog: https://gioxx.org/tag/o365-powershell
    Twitter: @gioxx    

.LINK
    https://morgantechspace.com/2021/01/check-size-and-status-of-archive-mailbox-powershell.html
    https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/archive-mailboxes/archivestatus-set-none
#> 

$Today = Get-Date -format yyyyMMdd
$CSV = "C:\temp\$($Today)_MailboxSize.csv"

<#
# Old sourcecode, to be removed in the future
Get-Mailbox -ResultSize Unlimited |
Select-Object DisplayName,
servername,database,
RecipientTypeDetails,PrimarySmtpAddress,
@{Name='TotalItemSize(GB)'; expression={[math]::Round((((Get-MailboxStatistics $_.PrimarySmtpAddress).TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},
IssueWarningQuota, ProhibitSendQuota |
Export-Csv $CSV -Append -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Invoke-Item $CSV
#>

Set-Variable ProgressPreference Continue
$Today = Get-Date -format yyyyMMdd
$CSV = "C:\temp\$($Today)_MailboxSize.csv"
$Result=@()
$ProcessedCount = 0
$Mailboxes = Get-Mailbox -ResultSize Unlimited
$TotalMailboxes = $Mailboxes.Count

$Mailboxes | Foreach-Object {
    $ProcessedCount++
    $Mbox = $_
    $Size = $null
    $ArchiveSize = $null
    Write-Progress -Activity "Processing $Mbox" -Status "$ProcessedCount out of $TotalMailboxes completed" -PercentComplete (($ProcessedCount / $TotalMailboxes) * 100)

    if ( $Mbox.ArchiveDatabase -ne $null ) {
        $MailboxArchiveSize = Get-MailboxStatistics $Mbox.UserPrincipalName -Archive
        if ( $MailboxArchiveSize.TotalItemSize -ne $null ) {
         $ArchiveSize = [math]::Round(($MailboxArchiveSize.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1GB),2)
        } else {
         $ArchiveSize = 0
        }
    }

    $MailboxSize = [math]::Round((((Get-MailboxStatistics $Mbox.UserPrincipalName).TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)

    $Result += New-Object -TypeName PSObject -Property $([ordered]@{ 
     UserName = $Mbox.DisplayName
     ServerName = $Mbox.ServerName
     Database = $Mbox.Database
     RecipientTypeDetails = $Mbox.RecipientTypeDetails
     PrimarySmtpAddress = $Mbox.PrimarySmtpAddress
     MailboxSizeInGB = $MailboxSize
     IssueWarningQuota = $Mbox.IssueWarningQuota
     ProhibitSendQuota = $Mbox.ProhibitSendQuota
     ArchiveDatabase = $Mbox.ArchiveDatabase
     ArchiveName = $Mbox.ArchiveName
     ArchiveState = $Mbox.ArchiveState
     ArchiveMailboxSizeInGB = $ArchiveSize
     ArchiveWarningQuota= if ( $Mbox.ArchiveDatabase -ne $null ) { $Mbox.ArchiveWarningQuota } else { $null } 
     ArchiveQuota = if ( $Mbox.ArchiveDatabase -ne $null ) { $Mbox.ArchiveQuota } else { $null } 
     AutoExpandingArchiveEnabled = $Mbox.AutoExpandingArchiveEnabled
    })
}
$Result | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"