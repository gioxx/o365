<#
	OFFICE 365: "User License Report"
	---------------------------------------------------------------------------------------------------
	Autore originale:	Kombaiah Murugan (8 feb. 2016)
	URL originale:		http://www.morgantechspace.com/2016/02/get-all-licensed-office-365-users-with-powershell.html
	
	Modifiche:			GSolone
	Versione:			0.4
	Utilizzo:			.\Export-MsolAccountSku.ps1
						(opzionale, posizione CSV) .\Export-MsolAccountSku.ps1 -CSV C:\Licenze.csv
						(opzionale, dominio da filtrare) .\Export-MsolAccountSku.ps1 -domain contoso.com
	Info:				http://gioxx.org/tag/o365-powershell
	Ultima modifica:	10-02-2017
	Modifiche:
		0.4- corretto errore nel "-f" quando si specifica un dominio di ricerca. Forzo un nome del CSV diverso se si specifica il dominio di ricerca.
		0.3- ho aggiunto il parametro UserPrincipalName all'estrazione, così da mostrare anche l'indirizzo di posta principale (generalmente corrispondente proprio a UserPrincipalName).
		0.2- includo la possibilità di filtrare un singolo dominio da riga di comando.
#>

# Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $CSV,
	[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Domain
)
	
# Main
""
Write-Host "        Office 365: User License Report" -foregroundcolor "Green"
Write-Host "        ------------------------------------------"
Write-Host "         Lo script crea un report delle licenze assegnate agli utenti" -f "White"
Write-Host "         configurati sul server Exchange, salvando i risultati su un file CSV" -f "White"
Write-Host "         '" -f "White" -nonewline; Write-Host $CSV -f "Green" -nonewline; Write-Host "'" -f "White"
Write-Host "         (rilancia lo script con parametro -CSV PERCORSOFILE.CSV per modificare)." -f "White"
Write-Host "         È possibile specificare un singolo dominio di ricerca ed esportazione da riga di comando." -f "White"
Write-Host "         (rilancia lo script con parametro -domain contoso.com per filtrare)." -f "White"
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

""
Write-Progress -Activity "Download dati da Exchange" -Status "Scarico i dati relativi alle licenze, attendi."
""

if ([string]::IsNullOrEmpty($Domain)) {
		# Dominio non specificato, analizzo tutti gli utenti del server Exchange
		$users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" }
	} else {
		# Dominio specificato, lo esporto in maniera esclusiva
		# Se non specifico il CSV, il nome integra il dominio di ricerca
		if ([string]::IsNullOrEmpty($CSV)) {
			$CSV = "C:\temp\O365-User-License-Report_$($Domain).csv"
		}
		""
		Write-Host "         Dominio specificato: " -nonewline; Write-Host "$Domain" -f "Yellow"
		Write-Host "         CSV di destinazione: " -nonewline; Write-Host "$CSV" -f "Yellow"
		$users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" -and $_.UserPrincipalName -like "*@" + $Domain }
	}
	
# Se il CSV non è stato precedentemente specificato, utilizzo posizione e nome di default	
if ([string]::IsNullOrEmpty($CSV)) {
		$CSV = "C:\temp\O365-User-License-Report.csv"
	}

$users | Foreach-Object {
	$licenseDetail = ''
	$licenses = ''
	if( $_.licenses -ne $null ) {
		
		# Licenze (piani Office e applicazioni aggiuntive)
		ForEach ( $license in $_.licenses ) {
			switch -wildcard ($($license.Accountskuid.tostring())) {
				'*POWER_BI_STANDALONE' { $licName = 'POWER BI STANDALONE' }
				'*POWER_BI_STANDARD' { $licName = 'POWER BI Standard' }
				'*POWER_BI_PRO' { $licName = 'POWER BI Pro' }
				'*DESKLESSPACK' { $licName = 'O365 Plan K1' }
				'*DESKLESSWOFFPACK' { $licName = 'O365 Plan K2' }
				'*CRMSTANDARD' { $licName = 'CRM Online' }
				'*STANDARD_B_PILOT' { $licName = 'O365 Small Business Preview' }
				'*O365_BUSINESS_PREMIUM' { $licName = 'O365 Business Premium' }
				'*ENTERPRISEPACK_B_PILOT' { $licName = 'O365 Enterprise Preview' }
				
				'*STANDARDPACK_STUDENT' { $licName = 'Office 365 (Plan A1) for Students' }
				'*STANDARDWOFFPACKPACK_STUDENT' { $licName = 'Office 365 (Plan A2) for Students' }
				'*ENTERPRISEPACK_STUDENT' { $licName = 'Office 365 (Plan A3) for Students' }
				'*ENTERPRISEWITHSCAL_STUDENT' { $licName = 'Office 365 (Plan A4) for Students' }
				'*STANDARDPACK_FACULTY' { $licName = 'Office 365 (Plan A1) for Faculty' }
				'*STANDARDWOFFPACKPACK_FACULTY' { $licName = 'Office 365 (Plan A2) for Faculty' }
				'*ENTERPRISEPACK_FACULTY' { $licName = 'Office 365 (Plan A3) for Faculty' }
				'*ENTERPRISEWITHSCAL_FACULTY' { $licName = 'Office 365 (Plan A4) for Faculty' }
			   
				'*LITEPACK' { $licName = 'O365 Plan P1' }
				'*EXCHANGESTANDARD' { $licName = 'O365 Plan P1 (Online Only)' }
				'*STANDARDPACK' { $licName = 'O365 Plan E1' }
				'*STANDARDWOFFPACK' { $licName = 'O365 Plan E2' }
				'*ENTERPRISEPACK' { $licName = 'O365 Plan E3' }
				'*ENTERPRISEPACKLRG' { $licName = 'O365 Plan E3' }
				'*ENTERPRISEWITHSCAL' { $licName = 'O365 Plan E4' }
				'*VISIOCLIENT' { $licName = 'Visio Pro 2016' }
				'*PROJECTCLIENT' { $licName = 'Project Pro 2016' }
				
				default { $licName = $license.Accountskuid.tostring() }
			}
			if ( $licenses ) { $licenses = ($licenses + ',' + $licName) } else { $licenses = $licName }
			#$licenses = $licName
			
			# Servizi
			ForEach ($row in $($license.servicestatus)) {
				if($row.ProvisioningStatus -ne 'Disabled') {
					switch -wildcard ($($row.ServicePlan.servicename)) {
						'EXC*' { $thisLicence = 'Exchange Online' }  
						'LYN*' { $thisLicence = 'Skype for Business' } 
						'SHA*' { $thisLicence = 'Sharepoint Online' }       
						default { $thisLicence = $row.ServicePlan.servicename }
					}
					if ($licenseDetail) { $licenseDetail = ($licenseDetail + ',' + $thisLicence) } else { $licenseDetail = $thisLicence }
				}
			}
		}
	}

	New-Object -TypeName PSObject -Property @{
		UserName=$_.DisplayName
		UserPrincipalName=$_.UserPrincipalName
		IsLicensed=$_.IsLicensed
		Licenses=$licenses
		<# 	Escludo i servizi di base del tenant. 
			Per includerli nuovamente occorre togliere il commento alla riga di seguito #>
		LicenseDetails=$licenseDetail
	}
<# 	Il Select finale non tiene conto dei servizi di base del tenant. 
	Per includerli nuovamente occorre togliere il commento alla riga di seguito e commentare il Select successivo,
	mantenendo però l'Export su CSV #>
#} | Select UserName,IsLicensed,Licenses,LicenseDetails | Export-CSV $CSV -NoTypeInformation -Encoding UTF8
} | Select UserName,UserPrincipalName,IsLicensed,Licenses | Export-CSV $CSV -NoTypeInformation -Encoding UTF8

Write-Host "Done." -f "Green"
""

# Chiedo se visualizzare i risultati esportati nel file CSV
$message = "Devo aprire il file CSV $CSV ?"
$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Apri il file ora"
$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non aprire il file adesso"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
$AproCSV = $host.ui.PromptForChoice("", $message, $options, 1)
if ($AproCSV -eq 0) { Invoke-Item $CSV }
""