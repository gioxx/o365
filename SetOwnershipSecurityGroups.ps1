############################################################################################################################
# OFFICE 365: Set multiple Ownership for all Security Groups
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1 rev2
# Utilizzo:				.\SetOwnershipSecurityGroups.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		10-03-2016
# Modifiche:
#	0.1 rev2- aggiunta funzione di Pausa per evitare di intercettare il tasto CTRL.
#	0.1 rev1- modifiche minori
############################################################################################################################

#Main
Function Main {


# Modificare la variabile se necessario, includendo ulteriori indirizzi tra le virgolette e spaziati da una virgola
$GroupOwners = "admin01@domain.tld", "admin02@domain.tld"


	""
	Write-Host "        Office 365: Set multiple Ownership for all Security Groups" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          Lo script cerca tutti i gruppi di sicurezza" -foregroundcolor "Cyan"
	Write-Host "          attualmente presenti sul server che hanno nome" -foregroundcolor "Cyan"
	Write-Host "          'Security - *' e riapplica l'ownership secondo" -foregroundcolor "Cyan"
	Write-Host "          gli utenti specificati nella variabile all'interno" -foregroundcolor "Cyan"
	Write-Host "          dello script stesso." -foregroundcolor "Cyan"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	""
	Write-Host "		Attuali owners dichiarati nello script: $GroupOwners"
	Write-Host "		Se non sono corretti termina questo script e modificalo inserendo i giusti riferimenti"
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
		Write-Host "Owners attualmente impostati nella variabile e che verranno riapplicati:" -foregroundcolor "yellow"
		$GroupOwners
		""
		Write-Host "Ricerco i gruppi di sicurezza presenti sul server Exchange, attendi." -foregroundcolor "yellow"
		$RicercaGruppi = Get-MsolGroup | where-object { $_.DisplayName -like "Security - *"}
		Write-Host "Done. Questi sono i gruppi di sicurezza attualmente presenti e trovati sul server:" -foregroundcolor "green"
		$RicercaGruppi | FT DisplayName,EmailAddress
		""
		Write-Host "Done. Applico l'ownership a tutti i gruppi, attendi." -foregroundcolor "green"
		$RicercaGruppi | ForEach-Object {Set-DistributionGroup $_.EmailAddress -ManagedBy $GroupOwners -BypassSecurityGroupManagerCheck}
		""
		Write-Host "Done." -foregroundcolor "green"
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
		write-host $error[0]
		return ""
	}
	
}

# Start script
. Main