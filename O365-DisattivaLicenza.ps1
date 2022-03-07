<#
OFFICE 365: Disable Office 365 License and Rearm
---------------------------------------------------------------------------------------------------
Modifiche:				GSolone
Versione:					0.1
Utilizzo:					.\O365-DisattivaLicenza.ps1
Info:							https://gioxx.org/tag/o365-powershell
Ultima modifica:	12-09-2017
Fonti utilizzate:	https://stackoverflow.com/questions/2988880/extricate-a-substring-using-powershell
Modifiche:
#>

# Cerco sulla macchina (in ordine) Office 2016 x64 / x86, poi -se non lo trovo- cerco Office 2013 x64 / x86
$OFC16x86 = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"
$OFC16x64 = "${env:ProgramFiles}\Microsoft Office\Office16\OSPP.VBS"
$OFC15x86 = "${env:ProgramFiles(x86)}\Microsoft Office\Office15\OSPP.VBS"
$OFC15x64 = "${env:ProgramFiles}\Microsoft Office\Office15\OSPP.VBS"

if (Test-Path -path $OFC16x64) {
	Write-Host "Trovato Office 2016 x64, licenze attive:";
	$license = cscript "$OFC16x64" /dstatus
} elseif (Test-Path -path $OFC16x86) {
	Write-Host "Trovato Office 2016 x86, licenze attive:";
	$license = cscript "$OFC16x86" /dstatus
} else {
	if (Test-Path -path $OFC15x64) {
		Write-Host "Trovato Office 2013 x64, licenze attive:";
		$license = cscript "$OFC15x64" /dstatus
	} elseif (Test-Path -path $OFC15x86) {
		Write-Host "Trovato Office 2013 x86, licenze attive:";
		$license = cscript "$OFC15x86" /dstatus
	}
}

$unpkey = $license -match "Last 5 characters of installed product key: (?<content>.*)"
$unpkey
$prodkey = $unpkey[0]
$prodkey = $prodkey.Substring(44,5)

if (Test-Path -path $OFC16x64) {
	""; Write-Host "Rimuovo licenza Office 2016 x64";
	cscript "$OFC16x64" /unpkey:$prodkey
	cscript "$OFC16x64" /rearm
} elseif (Test-Path -path $OFC16x86) {
	""; Write-Host "Rimuovo licenza Office 2016 x86";
	cscript "$OFC16x86" /unpkey:$prodkey
	cscript "$OFC16x86" /rearm
} else {
	if (Test-Path -path $OFC15x64) {
		""; Write-Host "Rimuovo licenza Office 2013 x64";
		cscript "$OFC15x64" /unpkey:$prodkey
		cscript "$OFC15x64" /rearm
	} elseif (Test-Path -path $OFC15x86) {
		""; Write-Host "Rimuovo licenza Office 2013 x86";
		cscript "$OFC15x86" /unpkey:$prodkey
		cscript "$OFC15x86" /rearm
	}
}
