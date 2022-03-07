<#	O365 PShell Snippet:	Export Quarantined Message
Autore (ver.-mod.):				GSolone (0.6 ult.mod. 29/10/20)
Utilizzo:									.\ExportQuarantinedMessage.ps1 55f33732-c398-e309-46a7-25202e43ae6a@contoso.com
Info:											https://gioxx.org/tag/o365-powershell
#>
Param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)][string] $MessageID)
if (-not($MessageID.StartsWith('<'))) { $MessageID = '<' + $MessageID }
if (-not($MessageID.EndsWith('>'))) { $MessageID += '>' }
$e = Get-QuarantineMessage -MessageId $($MessageID) | Export-QuarantineMessage; $bytes = [Convert]::FromBase64String($e.eml); [IO.File]::WriteAllBytes("C:\Temp\QuarantinedMessage.eml", $bytes)
Invoke-Item C:\Temp\QuarantinedMessage.eml
Start-Sleep -s 3
Remove-Item C:\Temp\QuarantinedMessage.eml
$message = "Devo rilasciare il messaggio a tutti i destinatari?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Rilascia il messaggio."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non rilasciare il messaggio."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice("", $message, $options, 0)
if ($result -eq 0) {
	Get-QuarantineMessage -MessageId $($MessageID) | Release-QuarantineMessage -ReleaseToAll
}
