############################################################################################################################
# OFFICE 365: Get Cluttered Mailboxes
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\List-ClutteredMailboxes.ps1
#						opzionale, output su CSV: .\List-ClutteredMailboxes.ps1 C:\temp\Clutter.csv
#						opzionale, numero mailbox analizzate: .\List-ClutteredMailboxes.ps1 -Count 10
#						opzionale, entrambi i parametri: .\List-ClutteredMailboxes.ps1 C:\temp\Clutter.csv 10
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		14-10-2015
# Modifiche:
#	0.2-	accetta output su CSV (esempio: .\List-ClutteredMailboxes.ps1 C:\temp\Clutter.csv) e numero di caselle da analizzare (esempio: .\List-ClutteredMailboxes.ps1 -Count 10)
#	0.1 rev2- modifiche minori
#	0.1 rev1- migliorato output dello script
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $CSV, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Count 
)

#Main
Function Main {
	
	""
	Write-Host "        Office 365: Get Cluttered Mailboxes" -f "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -f "red"
	Write-Host "          L'operazione può richiedere MOLTO tempo, dipende dal numero di utenti" -f "red"
	Write-Host "          da verificare e modificare all'interno della Directory, porta pazienza!" -f "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	""
	Write-Host "		Premi un tasto qualsiasi per continuare..."
	[void][System.Console]::ReadKey($true)
	
	try
	{
		""
		Write-Host "		A long time left, grab a Snickers!" -f yellow
		Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco tutte le caselle registrate nel sistema..."
		
		if ([string]::IsNullOrEmpty($CSV) -eq $true) {
		#CSV non dichiarato, output a video
			""
			Write-Host "		NESSUN CSV SPECIFICATO" -f "red"
			
			if ([string]::IsNullOrEmpty($Count) -eq $true) { 
				$Mailboxes = Get-Mailbox -ResultSize Unlimited
			} else {
				Write-Host "		Numero mailbox da analizzare: $Count" -f "green"
				""
				$Mailboxes = Get-Mailbox -ResultSize $Count
				""
			}
			
			Write-Host "Hanno la funzione Messaggi Secondari attiva: " -f yellow
			$Mailboxes | Foreach {
				$DN = $_.WindowsEmailAddress
				Write-Progress -Activity "Download dati da Exchange" -Status "Analizzo stato clutter di $DN" -PercentComplete (($i / $Mailboxes.count)*100)
				$StatoClutter = Get-Clutter -Identity $DN | Select -ExpandProperty isEnabled
				if ( $StatoClutter -eq "True" ) {
					Write-Host " - " $DN
				}
			}
		} else {
		#CSV dichiarato, output in file
			""
			Write-Host "		File CSV di output: $CSV" -f "green"
			
			if ([string]::IsNullOrEmpty($Count) -eq $true) { 
				$Mailboxes = Get-Mailbox -ResultSize Unlimited
			} else {
				Write-Host "		Numero mailbox da analizzare: $Count" -f "green"
				""
				$Mailboxes = Get-Mailbox -ResultSize $Count
				""
			}
			
			Write-Host "Hanno la funzione Messaggi Secondari attiva: " -f yellow
			$Mailboxes | Foreach {
				$DN = $_.WindowsEmailAddress
				Write-Progress -Activity "Download dati da Exchange" -Status "Analizzo stato clutter di $DN" -PercentComplete (($i / $Mailboxes.count)*100)
				$StatoClutter = Get-Clutter -Identity $DN | Select -ExpandProperty isEnabled
				if ( $StatoClutter -eq "True" ) {
					Write-Host " - " $DN
					Out-File -FilePath $CSV -InputObject "$DN" -Encoding UTF8 -append
				}
			}
			Invoke-Item $CSV
		}
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -f "red"
		write-host $error[0]
		return ""
	}
	
}

# Start script
. Main