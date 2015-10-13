############################################################################################################################
# OFFICE 365: Disable Skype for Business (Bulk, All Users)
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\DisableSkypeForBusiness.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		19-05-2015
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Disable Skype for Business for all users" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "        Lo script ricerca e disabilita Skype for Business su tutte le utenze" -f "white"
	Write-Host "        che hanno licenza Enterprise (E1, E3), sono escluse le Online (P1, P2)" -f "white"
	Write-Host "        perché non hanno possibilità nativa di sfruttare il servizio" -f "white"
	""
	Write-Host "		Premi un tasto qualsiasi per continuare..."
	[void][System.Console]::ReadKey($true)
	
	try
	{
		""; "";
		Write-Progress -Activity "Scambio dati con Exchange" -Status "Analisi pacchetti licenza disponibili..."
		
		Write-Host "Stato attuale delle licenze disponibili su Exchange" -f "yellow"
		Get-MsolAccountSku | ft AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits -AutoSize | out-string
		
		Write-Progress -Activity "Scambio dati con Exchange" -Status "Assegnazione nuova licenza E1"

		Write-Host "Elenco servizi legati a licenza E1" -f "yellow"
		Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "STANDARDPACK"} | ForEach-Object {$_.ServiceStatus} | out-string

		Write-Host "Elenco utenze registrate con licenza E1" -f "yellow"
		$O365E1Licences = New-MsolLicenseOptions -AccountSkuId messita:STANDARDPACK -DisabledPlans MCOSTANDARD
		Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "messita:STANDARDPACK"} | ft UserPrincipalName, DisplayName, isLicensed -AutoSize | out-string

		Write-Host "Disabilito Skype for Business sulle licenze E1 ... " -f "yellow" -nonewline
		Get-MsolUser -all |  where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "messita:STANDARDPACK"} | Set-MsolUserLicense -LicenseOptions $O365E1Licences
		Write-Host " fatto" -f "green" -nonewline; Write-Host "."

		Write-Progress -Activity "Scambio dati con Exchange" -Status "Assegnazione nuova licenza E3"
		
		""; Write-Host "Elenco servizi legati a licenza E3" -f "yellow"
		Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPACK"} | ForEach-Object {$_.ServiceStatus} | out-string
		
		Write-Host "Elenco utenze registrate con licenza E3" -f "yellow"
		$O365E3Licences = New-MsolLicenseOptions -AccountSkuId messita:ENTERPRISEPACK -DisabledPlans MCOSTANDARD
		Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "messita:ENTERPRISEPACK"} | ft UserPrincipalName, DisplayName, isLicensed -AutoSize | out-string

		Write-Host "Disabilito Skype for Business sulle licenze E3 ... " -f "yellow" -nonewline
		Get-MsolUser -all |  where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "messita:ENTERPRISEPACK"} | Set-MsolUserLicense -LicenseOptions $O365E3Licences
		Write-Host " fatto" -f "green" -nonewline; Write-Host "."
		""; "";
		Write-Host "Verifica almeno un'utenza per ciascuna licenza sulla GUI di Exchange. Vai in"
		Write-Host "modifica della licenza, espandi i pacchetti assegnati dalla licenza e verifica che"
		Write-Host "non compaia il segno di spunta in corrispondenza di Skype for Business."
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