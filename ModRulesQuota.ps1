############################################################################################################################
# OFFICE 365: Set Rules Quota to 128Kb
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\ModRulesQuota.ps1
# Info:					http://gioxx.org/tag/o365-powershell
#						http://social.technet.microsoft.com/Forums/exchange/en-US/1b5d268f-014b-4914-ad6a-d10165262d24/rules-do-not-match
# Ultima modifica:		22-07-2014
# Modifiche:			-
############################################################################################################################

#DEBUG:
# $FPath = "C:\temp\"
# $fileOut = $FPath + "RulesQuota_LOG.txt"
# Get-Mailbox | foreach { 
#
# Set-Mailbox $_.UserPrincipalName -RulesQuota 128KB
# Write-Host "Set Rules Quota to 128KB: " $_.UserPrincipalName
#
# }|Out-File -FilePath $fileOut

""
Write-Host "        Office 365: Set Rules Quota to 128Kb" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -foregroundcolor "red"
Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"
$QuotaUser = Read-Host "Utente (esempio: info@domain.tld)"

try
	{
		""
		Write-Host " $QuotaUser : situazione attuale" -foregroundcolor "yellow"
		Get-Mailbox $QuotaUser | FT UserPrincipalName,RulesQuota
		Set-Mailbox $QuotaUser -RulesQuota 128KB
		Write-Host " $QuotaUser : nuova situazione" -foregroundcolor "green"
		Get-Mailbox $QuotaUser | FT UserPrincipalName,RulesQuota
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
		write-host $error[0]
		return ""
	}