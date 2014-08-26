############################################################################################################################
# OFFICE 365: Reset User Password from PowerShell
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.2
# Utilizzo:				.\ResetPassword.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		08-04-2014 (15072014-rev2)
# Modifiche:
#	0.2- correzioni minori.
############################################################################################################################

""
Write-Host "        Office 365: Single Password Reset" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -foregroundcolor "red"
Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"
$ResetUser = Read-Host "Utente (esempio: info@domain.tld)"

try
	{
		Set-MsolUserPassword -UserPrincipalName $ResetUser -NewPassword "Office2013"
		Set-MsolUser -UserPrincipalName $ResetUser -PasswordNeverExpires $true
		#DEBUG Get-MSOLUser -UserPrincipalName $ResetUser | Select PasswordNeverExpires
		""
		Write-Host "OK: $ResetUser" -foregroundcolor "green"
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
		write-host $error[0]
		return ""
	}