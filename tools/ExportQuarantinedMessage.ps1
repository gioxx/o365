<#	O365 PShell Snippet:	Export Quarantined Message
	Autore (ver.-mod.):		GSolone (0.2 ult.mod. 18/9/20)
	Utilizzo:				.\ExportQuarantinedMessage.ps1 dd175158-3973-40da-7e68-08d847c9c1b7\c9b12f14-728a-d306-1d58-2d84051c2294
	Info:					https://gioxx.org/tag/o365-powershell
#>
Param([Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)][string] $MessageID)
$idnt = Get-QuarantineMessage -MessageId $($MessageID) | Select -ExpandProperty Identity
$e = Export-QuarantineMessage -Identity $($idnt); $bytes = [Convert]::FromBase64String($e.eml); [IO.File]::WriteAllBytes("C:\Temp\QuarantinedMessage.eml", $bytes)
Invoke-Item C:\Temp\QuarantinedMessage.eml
Start-Sleep -s 3
Remove-Item C:\Temp\QuarantinedMessage.eml