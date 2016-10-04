#####################################################################
# O365 PShell Snippet:	Get Mailbox Used Disk						#
# Autore (ver.-mod.):	GSolone (0.1 ult.mod. 20/9/16)				#
# Utilizzo:				.\mboxUsedDisk.ps1 user@contoso.com			#
# Info:					http://gioxx.org/tag/o365-powershell		#
#####################################################################

<#
	Controllo i seguenti valori per l'occupazione attuale su disco e per le statistiche:
		-	DisplayName
		-	LastLogonTime
		-	TotalItemSize
		-	ItemCount
		-	TotalDeletedItemSize

#>

# Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $SourceMailbox
)

# Dettagli occupazione su disco
""
Write-Host "Verifico ultimo login e occupazione disco della casella $SourceMailbox " -f "Green"
Get-MailboxStatistics $SourceMailbox | ft DisplayName, LastLogonTime, ItemCount, TotalItemSize, TotalDeletedItemSize
# Quote disco
$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Vedi quote disco"
$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non vedere quote disco"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
$QuoteDisco = $host.ui.PromptForChoice("Vuoi vedere le quote disco?", $message, $options, 1)
if ($QuoteDisco -eq 0) { 
	Write-Host "Ulteriori statistiche di $SourceMailbox " -f "yellow"
	Get-MailboxStatistics $SourceMailbox | ft DatabaseIssueWarningQuota, DatabaseProhibitSendQuota, DatabaseProhibitSendReceiveQuota	
}
""