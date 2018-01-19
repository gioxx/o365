<#
	OFFICE 365: Disable Skype for Business (Bulk, All Users)
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.3
	Utilizzo:			.\DisableSkypeForBusiness.ps1 -Company AccountSkuId*
						*(	How to obtain AccountSkuId:
							From PowerShell execute "Get-MsolAccountSku | ft AccountSkuId"
								AccountSkuId
								------------
								CONTOSO:VISIOCLIENT
								CONTOSO:POWER_BI_PRO
								CONTOSO:ENTERPRISEPACK
								CONTOSO:FLOW_FREE
								CONTOSO:MCOEV
								CONTOSO:EXCHANGESTANDARD
								CONTOSO:DYN365_ENTERPRISE_PLAN1
								CONTOSO:POWER_BI_STANDARD
								CONTOSO:PROJECTPROFESSIONAL
								CONTOSO:DYN365_ENTERPRISE_TEAM_MEMBERS
								CONTOSO:STANDARDPACK
							AccountSkuId = "CONTOSO", then:
							.\DisableSkypeForBusiness.ps1 -Company CONTOSO
						)
	Info:				http://gioxx.org/tag/o365-powershell
	Ultima modifica:	19-01-2018
	Modifiche:			
		0.3- corretto un problema relativo alla ricerca degli utenti, non venivano più trovati. Comando sostituito con equivalente funzionante, vedi https://community.spiceworks.com/topic/post/7275311.
		0.2- introdotto il parametro Company per modificare correttamente i pacchetti di licenza aziendali (l'identificativo è unico per ciascun cliente Office 365), di conseguenza modificato il metodo per comporre la nuova licenza. Ho poi modificato il messaggio finale di controllo dell'utenza a campione con l'apertura della console di amministrazione Lync tramite il browser predefinito di sistema.
		0.1 rev1- aggiunta funzione di Pausa per evitare di intercettare il tasto CTRL.
#>

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $Company 
)

#Main
Function Main {

	""
	Write-Host "        Office 365: Disable Skype for Business for all users" -foregroundcolor "Green"
	Write-Host "        ------------------------------------------"
	Write-Host "        Lo script ricerca e disabilita Skype for Business su tutte le utenze" -f "white"
	Write-Host "        che hanno licenza Enterprise (E1, E3), sono escluse le Online (P1, P2)" -f "white"
	Write-Host "        perché non hanno possibilità nativa di sfruttare il servizio" -f "white"
	""
	
	Function Pause($M="Premi un tasto per continuare (CTRL+C per annullare)") {
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
		""; "";
		Write-Progress -Activity "Scambio dati con Exchange" -Status "Analisi pacchetti licenza disponibili..."
		Write-Host "Stato attuale delle licenze disponibili su Exchange" -f "Yellow"
		Get-MsolAccountSku | ft AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits -AutoSize | Out-String
		
<# ----	MODIFICA LICENZE E1 --------------------------------------------------------------------------------------------------------------- #>

		Write-Progress -Activity "Scambio dati con Exchange" -Status "Assegnazione nuova licenza E1"
		Write-Host "Elenco servizi legati a licenza E1" -f "Yellow"
		Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "STANDARDPACK"} | ForEach-Object {$_.ServiceStatus} | Out-String

		Write-Host "Elenco utenze registrate con licenza E1" -f "Yellow"
		$command = '$O365E1Licences = New-MsolLicenseOptions -AccountSkuId ' + $Company + ':STANDARDPACK -DisabledPlans MCOSTANDARD'
		iex $command
		Get-MsolUser -All | where {($_.islicensed -eq $true) -and ($_.licenses.accountskuid -like "*STANDARDPACK*")} | ft UserPrincipalName, DisplayName, isLicensed -AutoSize | Out-String

		Write-Host "Disabilito Skype for Business sulle licenze E1 ... " -f "Yellow" -nonewline
		Get-MsolUser -All |  where {($_.islicensed -eq $true) -and ($_.licenses.accountskuid -like "*STANDARDPACK*")} | Set-MsolUserLicense -LicenseOptions $O365E1Licences
		Write-Host " fatto" -f "Green" -nonewline; Write-Host "."
		
<# ----	MODIFICA LICENZE E3 --------------------------------------------------------------------------------------------------------------- #>

		Write-Progress -Activity "Scambio dati con Exchange" -Status "Assegnazione nuova licenza E3"
		""; Write-Host "Elenco servizi legati a licenza E3" -f "Yellow"
		Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPACK"} | ForEach-Object {$_.ServiceStatus} | Out-String
		
		Write-Host "Elenco utenze registrate con licenza E3" -f "Yellow"
		$command = '$O365E3Licences = New-MsolLicenseOptions -AccountSkuId ' + $Company + ':ENTERPRISEPACK -DisabledPlans MCOSTANDARD'
		iex $command
		#Debug, vedi com'era:
		#Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "$($Company):ENTERPRISEPACK"} | ft UserPrincipalName, DisplayName, isLicensed -AutoSize | Out-String
		Get-MsolUser -All | where {($_.islicensed -eq $true) -and ($_.licenses.accountskuid -like "*ENTERPRISEPACK*")} | ft UserPrincipalName, DisplayName, isLicensed -AutoSize | Out-String

		Write-Host "Disabilito Skype for Business sulle licenze E3 ... " -f "Yellow" -nonewline
		Get-MsolUser -All |  where {($_.islicensed -eq $true) -and ($_.licenses.accountskuid -like "*ENTERPRISEPACK*")} | Set-MsolUserLicense -LicenseOptions $O365E3Licences
		Write-Host " fatto" -f "Green" -nonewline; Write-Host "."
		
		""; "";
		Write-Host "Avvio il browser per verificare se gli utenti compaiono in console ..."
		Start-Process -FilePath "https://admin0e.online.lync.com/LSCP/Users.aspx"
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