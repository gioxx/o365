############################################################################################################################
# OFFICE 365: Set Password to Never Expire (BULK)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetPasswordNeverExpireBulk.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		15-07-2014
# Modifiche:			-
############################################################################################################################

#$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $PWord
#Import-Module MSOnline
#Connect-MsolService -Credential $Credential

#Main
Function Main {

	""
	Write-Host "        Office 365: Set Password to Never Expire (BULK)" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          L'operazione può richiedere MOLTO tempo, dipende dal numero di utenti" -foregroundcolor "red"
	Write-Host "          da verificare e modificare all'interno della Directory, porta pazienza!" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$RicercaUser = Read-Host "Premi invio per dare inizio alle danze "
	
	try
	{
		#DEBUG NO-SCADENZA Get-MSOLUser -All | Set-MsolUser -PasswordNeverExpires $false
		Write-Host "A long time left, grab a Snickers!" -foregroundcolor "yellow"
		Get-MSOLUser -All | Set-MsolUser -PasswordNeverExpires $true
		""
		Get-MSOLUser -All | Select UserPrincipalName, PasswordNeverExpires > C:\temp\PasswordNeverExpires.txt
		Write-Host "Ho esportato la nuova lista di utenti con scadenza password in C:\temp\PasswordNeverExpires.txt" -foregroundcolor "green"
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