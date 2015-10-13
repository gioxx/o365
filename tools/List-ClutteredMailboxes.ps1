############################################################################################################################
# OFFICE 365: Get Cluttered Mailboxes
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1 rev2
# Utilizzo:				.\ClutteredMailboxes.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		13-10-2015
# Modifiche:
#	0.1 rev2- modifiche minori
#	0.1 rev1- migliorato output dello script
############################################################################################################################

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
		$Mailboxes = Get-Mailbox -ResultSize Unlimited
		""

		Write-Host "Hanno la funzione Messaggi Secondari attiva: " -f yellow
		$Mailboxes | Foreach {
			$DN = $_.WindowsEmailAddress
			Write-Progress -Activity "Download dati da Exchange" -Status "Analizzo stato clutter di $DN" -PercentComplete (($i / $Mailboxes.count)*100)
			$StatoClutter = Get-Clutter -Identity $DN | Select -ExpandProperty isEnabled
			if ( $StatoClutter -eq "True" ) {
				Write-Host " - " $DN
			}
		}
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