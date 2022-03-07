<#	O365 PShell Snippet:	 Export Personal Mailboxes
Autore (ver.-mod.):		     GSolone (0.1 ult.mod. 19/7/19)
Utilizzo:				           .\ExportPersonalMailboxes.ps1 contoso.com
Info:					             https://gioxx.org/tag/o365-powershell
#>
Param( [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)][string] $Domain )
$Users = Get-Recipient -ResultSize unlimited |
Where {$_.RecipientType -eq "UserMailbox" -and $_.ResourceType -ne "Room" -and $_.ResourceType -ne "Equipment" -and $_.RecipientTypeDetails -ne "SharedMailbox" -and $_.PrimarySmtpAddress -like "*@$($Domain)"}
$Users | ft DisplayName,Name,PrimarySmtpAddress -AutoSize
$Users | select DisplayName,Name,PrimarySmtpAddress | Export-Csv C:\temp\$($Domain).csv -NoTypeInformation -Delimiter ";"
$message = "Devo aprire il file CSV $($Domain).csv?"
$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Apri il file"
$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non aprire ora il file"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
$Excel = $host.ui.PromptForChoice("", $message, $options, 0)
if ($Excel -eq 0) { Invoke-Item C:\temp\$($Domain).csv }
""
