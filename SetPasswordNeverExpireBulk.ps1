############################################################################################################################
# OFFICE 365: Set Password to Never Expire (BULK)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1 rev2
# Utilizzo:				.\SetPasswordNeverExpireBulk.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		15-06-2016
# Modifiche:
#	0.1 rev2- modifiche minori
#	0.1 rev1- modifiche minori
############################################################################################################################

#$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $PWord
#Import-Module MSOnline
#Connect-MsolService -Credential $Credential

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
	Write-Host "		Premi un tasto qualsiasi per continuare..."
	[void][System.Console]::ReadKey($true)
	
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