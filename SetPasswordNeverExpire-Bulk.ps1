############################################################################################################################
# OFFICE 365: Set Password to Never Expire (BULK)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1 rev3
# Utilizzo:				.\SetPasswordNeverExpireBulk.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		10-03-2016
# Modifiche:
#	0.1 rev3- aggiunta funzione di Pausa per evitare di intercettare il tasto CTRL.
#	0.1 rev2- modifiche minori
#	0.1 rev1- modifiche minori
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Set Password to Never Expire (BULK)" -f Green
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -f Red
	Write-Host "          L'operazione può richiedere MOLTO tempo, dipende dal numero di utenti" -f Red
	Write-Host "          da verificare e modificare all'interno della Directory, porta pazienza!" -f Red
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
		#DEBUG NO-SCADENZA Get-MSOLUser -All | Set-MsolUser -PasswordNeverExpires $false
		Write-Host "A long time left, grab a Snickers!" -f Yellow
		Get-MSOLUser -All | Set-MsolUser -PasswordNeverExpires $true
		Get-MSOLUser -All | Select UserPrincipalName, PasswordNeverExpires > C:\temp\PasswordNeverExpires.txt
		""
		Write-Host "Ho esportato la nuova lista di utenti con scadenza password in C:\temp\PasswordNeverExpires.txt" -f Green
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -f Red
		write-host $error[0]
		return ""
	}
	
}

# Start script
. Main