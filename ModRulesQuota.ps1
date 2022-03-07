############################################################################################################################
# OFFICE 365: Set Rules Quota to 256 KB
#----------------------------------------------------------------------------------------------------------------
# Autore:							GSolone
# Versione:						0.2
# Utilizzo:						.\ModRulesQuota.ps1
# Info:								https://gioxx.org/tag/o365-powershell
#											http://social.technet.microsoft.com/Forums/exchange/en-US/1b5d268f-014b-4914-ad6a-d10165262d24/rules-do-not-match
# Ultima modifica:		30-09-2016
# Modifiche:
#		0.2- portata quota a 256K
############################################################################################################################

<#
DEBUG:
	$FPath = "C:\temp\"
	$fileOut = $FPath + "RulesQuota_LOG.txt"
	Get-Mailbox | foreach {
	 Set-Mailbox $_.UserPrincipalName -RulesQuota 256KB
	 Write-Host "Set Rules Quota to 128KB: " $_.UserPrincipalName
	}|Out-File -FilePath $fileOut
#>

""; Write-Host "        Office 365: Set Rules Quota to 256 KB" -f "Green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -f "Red"
Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -f "Red"
Write-Host "          nei dati richiesti qui di seguito" -f "Red"; "";
Write-Host "-------------------------------------------------------------------------------------------------"
$QuotaUser = Read-Host "Utente (esempio: info@contoso.com)"

try {
	""; Write-Host " $QuotaUser : situazione attuale" -f "yellow";
	Get-Mailbox $QuotaUser | FT UserPrincipalName,RulesQuota
	Set-Mailbox $QuotaUser -RulesQuota 256KB
	Write-Host " $QuotaUser : nuova situazione" -f "Green"
	Get-Mailbox $QuotaUser | FT UserPrincipalName,RulesQuota
	""
} catch {
	Write-Host "Errore nell'operazione, riprovare." -f "Red"
	write-host $error[0]
	return ""
}
