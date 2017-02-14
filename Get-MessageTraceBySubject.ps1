<#
	OFFICE 365: "Get-MessageTrace by Mail Subject"
	---------------------------------------------------------------------------------------------------
	Autore originale:	Joe Palarchio (25 feb. 2016)
	URL originale:		http://blogs.perficient.com/microsoft/2016/03/office-365-script-to-perform-message-trace-by-subject/
	Limitazioni:		Search query is limited to 5,000,000 entries
	
	Modifiche:			GSolone
	Versione:			0.1 (versione script originale: 1.1)
	Utilizzo:			.\Get-MessageTraceBySubject.ps1
						(OBBLIGATORIO, oggetto mail da filtrare) .\Get-MessageTraceBySubject.ps1 -Subject "Hello Mario!"
						(OBBLIGATORIO, numeri giorni di ricerca) .\Get-MessageTraceBySubject.ps1 -Days 5
						(opzionale, posizione file di log) .\Get-MessageTraceBySubject.ps1 -LogFile C:\MessageTrace.txt
	Info:				http://gioxx.org/tag/o365-powershell
	Ultima modifica:	14-02-2017
	Modifiche:			-
#>

Param(
    [Parameter(Mandatory=$True)]
        [int]$Days,
    [Parameter(Mandatory=$True)]
        [string]$Subject,
    [Parameter(Mandatory=$False)]
        [string]$LogFile
    )

[DateTime]$DateEnd = Get-Date -format "MM/dd/yyyy HH:mm"
[DateTime]$DateStart = $DateEnd.AddDays($Days * -1)

if ([string]::IsNullOrEmpty($LogFile)) {
	$DataAdesso = Get-Date -format "ddMMyyy-HHmm"
	$LogFile = "MessageTrace_$($DataAdesso).txt"
}

$FoundCount = 0
For($i = 1; $i -le 1000; $i++)  # Maximum allowed pages is 1000
{
    $Messages = Get-MessageTrace -StartDate $DateStart -EndDate $DateEnd -PageSize 5000 -Page $i
    If($Messages.count -gt 0)
    {
        $Status = $Messages[-1].Received.ToString("MM/dd/yyyy HH:mm") + " - " + $Messages[0].Received.ToString("MM/dd/yyyy HH:mm") + "  [" + ("{0:N0}" -f ($i*5000)) + " analizzati , " + $FoundCount + " trovati]"
        Write-Progress -activity "Ricerca messaggio in corso ..." -status $Status
        $Entries = $Messages | Where {$_.Subject -like $Subject} | Select Received, SenderAddress, RecipientAddress, Subject, Status, FromIP, Size, MessageId
        $Entries | Out-File -FilePath $LogFile -Encoding UTF8 -Append
        $FoundCount += $Entries.Count
    } else { Break }
}  

""
Write-Host $FoundCount -f "Yellow" -nonewline; Write-Host " voci trovate e tracciate in $($LogFile)"
""