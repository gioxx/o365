<#	O365 PShell Snippet:	Split CSS in multiple CSS with same header
Autore (ver.-mod.):				GSolone (0.1 ult.mod. 6/11/20)
Utilizzo:									.\ExportPersonalMailboxes.ps1 C:\mailboxes.csv
Info:											https://gioxx.org/tag/o365-powershell
Credits:									https://stephan.io/split-a-csv-in-windows-powershell/
													https://devblogs.microsoft.com/scripting/powertip-read-first-line-of-file-with-powershell/
													https://powershell.org/forums/topic/read-first-line-and-then-delete-it/
#>

Param( [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [string] $sourceCSV )
$header = Get-Content $sourceCSV -First 1
Set-Content $PSScriptRoot\temp.csv (Get-Content $sourceCSV | select -skip 1)
$i=0
Get-Content $PSScriptRoot\temp.csv -ReadCount 100 | %{
	$i++
	$header | Out-File C:\temp\split_$i.csv
	$_ | Out-File C:\temp\split_$i.csv -Append
}
del $PSScriptRoot\temp.csv
