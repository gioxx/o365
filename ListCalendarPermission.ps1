<#
	OFFICE 365: List Calendars User Access (CSV)
	-------------------------------------------------------------------------------------------------------------
	Autore:					GSolone
	Versione:				0.1
	Utilizzo:				.\ListCalendarPermission.ps1
	Info:					https://gioxx.org/tag/o365-powershell
	Ultima modifica:		05-11-2020
	Fonti utilizzate:		https://morgantechspace.com/2019/09/get-calendar-permissions-for-all-users-using-powershell.html
							https://itsallinthecode.com/exchange-powershell-get-calendar-folder-permissions-in-any-language/
							https://techcommunity.microsoft.com/t5/office-365/report-on-default-calendar-permissions-if-they-are-set-to/m-p/155060
							https://devblogs.microsoft.com/scripting/use-a-powershell-cmdlet-to-count-files-words-and-lines/
	Modifiche:
		-
	
	ATTENZIONE:
	Il CSV deve contenere in colonna 1 il DisplayName e in colonna 2 il PrimarySmtpAddress degli utenti da analizzare.
#>

$Result=@()
$allMailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object -Property Displayname,PrimarySMTPAddress
$totalMailboxes = $allMailboxes.Count
$i = 1 
$allMailboxes | ForEach-Object {
	$mailbox = $_
	Write-Progress -activity "Processing $($_.Displayname)" -status "$i out of $totalMailboxes completed"
	$calendarFolder = Get-MailboxFolderStatistics -Identity $_.PrimarySMTPAddress -FolderScope Calendar | Where-Object { $_.FolderType -eq 'Calendar'} | Select-Object Name,FolderId
	$folderPerms = Get-MailboxFolderPermission -Identity "$($_.PrimarySMTPAddress):$($calendarFolder.FolderId)" | where { $_.AccessRights -notlike "AvailabilityOnly" -and $_.AccessRights -notlike "None"} 
	$folderPerms | ForEach-Object {
		$Result += New-Object PSObject -property @{ 
		MailboxName = $mailbox.DisplayName
		User = $_.User
		Permissions = $_.AccessRights
		}
	}
	$i++
}
$Result | Select MailboxName, User, Permissions |
Export-CSV "C:\temp\CalendarPermissions.CSV" -NoTypeInformation -Encoding UTF8 -Delimiter ";"