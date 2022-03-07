<#
OFFICE 365: List Calendars User Access (CSV)
-------------------------------------------------------------------------------------------------------------
Autore:							GSolone
Versione:						0.2
Utilizzo:						.\ListCalendarPermission-CSV.ps1 -CSV C:\Utenti.csv
Info:								https://gioxx.org/tag/o365-powershell
Ultima modifica:		06-11-2020
Fonti utilizzate:		https://morgantechspace.com/2019/09/get-calendar-permissions-for-all-users-using-powershell.html
										https://itsallinthecode.com/exchange-powershell-get-calendar-folder-permissions-in-any-language/
										https://techcommunity.microsoft.com/t5/office-365/report-on-default-calendar-permissions-if-they-are-set-to/m-p/155060
										https://devblogs.microsoft.com/scripting/use-a-powershell-cmdlet-to-count-files-words-and-lines/
Modifiche:
		0.2- ritocco il modo di procurarmi la localizzazione della cartella "Calendar" perchè in molti casi il metodo iniziale portava a errori in PowerShell (saltando così la verifica dell'utente).

ATTENZIONE:
		Il CSV deve contenere in colonna 1 il DisplayName e in colonna 2 il PrimarySmtpAddress degli utenti da analizzare.
#>

Param( [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [string] $CSV )
$Result=@()
$totalMailboxes = Get-Content $CSV | Measure-Object -Line
$totalMailboxes = $totalMailboxes.Lines - 1
$i = 1
Import-CSV $CSV | ForEach-Object {
	$mailbox = $_
	Write-Progress -activity "Verifico $($_.Displayname)" -status "$i di $totalMailboxes completati"
	$calendarFolder = Get-MailboxFolderStatistics -Identity "$($_.PrimarySMTPAddress)" -FolderScope Calendar | Select Name -ExpandProperty Name -First 1
	$folderPerms = Get-MailboxFolderPermission -Identity "$($_.PrimarySMTPAddress):\$($calendarFolder)" | where { $_.AccessRights -notlike "AvailabilityOnly" -and $_.AccessRights -notlike "None"}
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
Export-CSV "C:\temp\CalendarPermissions.CSV" -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Append
