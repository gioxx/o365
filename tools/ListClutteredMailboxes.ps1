############################################################################################################################
# OFFICE 365: Get Cluttered Mailboxes
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2 rev1
# Utilizzo:				.\List-ClutteredMailboxes.ps1
#						opzionale, output su CSV: .\List-ClutteredMailboxes.ps1 C:\temp\Clutter.csv
#						opzionale, numero mailbox analizzate: .\List-ClutteredMailboxes.ps1 -Count 10
#						opzionale, entrambi i parametri: .\List-ClutteredMailboxes.ps1 C:\temp\Clutter.csv 10
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		10-03-2016
# Modifiche:
#	0.2 rev1- aggiunta funzione di Pausa per evitare di intercettare il tasto CTRL.
#	0.2- accetta output su CSV (esempio: .\List-ClutteredMailboxes.ps1 C:\temp\Clutter.csv) e numero di caselle da analizzare (esempio: .\List-ClutteredMailboxes.ps1 -Count 10)
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
	
	Function Pause($M="Premi un tasto continuare (CTRL+C per annullare)") {
		If($psISE) {
			$S=New-Object -ComObject "WScript.Shell"
			$B=$S.Popup("Fai clic su OK per continuare.",0,"In attesa dell'amministratore",0)
			Return
		}
		Write-Host -NoNewline $M;
		$I=16,17,18,20,91,92,93,144,145,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183;
		While($K.VirtualKeyCode -Eq $Null -Or $I -Contains $K.VirtualKeyCode) {
			$K=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}
	}
	Pause
	
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