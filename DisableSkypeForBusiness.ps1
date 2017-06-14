<#
	OFFICE 365: Disable Skype for Business (Bulk, All Users)
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.2
	Utilizzo:			.\DisableSkypeForBusiness.ps1 -Company CONTOSO (a CONTOSO dovrai sostituire l'identificativo della tua azienda per modificare correttamente i pacchetti di licenza)
	Info:				http://gioxx.org/tag/o365-powershell
	Ultima modifica:	11-04-2017
	Modifiche:			
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
	Write-Host "        Office 365: Disable Skype for Business for all users" -foregroundcolor "green"
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
		
		Write-Host "Stato attuale delle licenze disponibili su Exchange" -f "yellow"
		Get-MsolAccountSku | ft AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits -AutoSize | out-string
		
		Write-Progress -Activity "Scambio dati con Exchange" -Status "Assegnazione nuova licenza E1"
		
		################################################
		# MODIFICA LICENZE E1
		################################################

		Write-Host "Elenco servizi legati a licenza E1" -f "yellow"
		Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "STANDARDPACK"} | ForEach-Object {$_.ServiceStatus} | out-string

		Write-Host "Elenco utenze registrate con licenza E1" -f "yellow"
		$command = '$O365E1Licences = New-MsolLicenseOptions -AccountSkuId ' + $Company + ':STANDARDPACK -DisabledPlans MCOSTANDARD'
		iex $command
		Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "$($Company):STANDARDPACK"} | ft UserPrincipalName, DisplayName, isLicensed -AutoSize | out-string

		Write-Host "Disabilito Skype for Business sulle licenze E1 ... " -f "yellow" -nonewline
		Get-MsolUser -all |  where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "$($Company):STANDARDPACK"} | Set-MsolUserLicense -LicenseOptions $O365E1Licences
		Write-Host " fatto" -f "green" -nonewline; Write-Host "."
		
		################################################
		# MODIFICA LICENZE E3
		################################################

		Write-Progress -Activity "Scambio dati con Exchange" -Status "Assegnazione nuova licenza E3"
		
		""; Write-Host "Elenco servizi legati a licenza E3" -f "yellow"
		Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPACK"} | ForEach-Object {$_.ServiceStatus} | out-string
		
		Write-Host "Elenco utenze registrate con licenza E3" -f "yellow"
		$command = '$O365E3Licences = New-MsolLicenseOptions -AccountSkuId ' + $Company + ':ENTERPRISEPACK -DisabledPlans MCOSTANDARD'
		iex $command
		Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "$($Company):ENTERPRISEPACK"} | ft UserPrincipalName, DisplayName, isLicensed -AutoSize | out-string

		Write-Host "Disabilito Skype for Business sulle licenze E3 ... " -f "yellow" -nonewline
		Get-MsolUser -all |  where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq "$($Company):ENTERPRISEPACK"} | Set-MsolUserLicense -LicenseOptions $O365E3Licences
		Write-Host " fatto" -f "green" -nonewline; Write-Host "."
		
		
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