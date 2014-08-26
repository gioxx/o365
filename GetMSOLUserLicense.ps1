############################################################################################################################
# OFFICE 365: Get MSOL User License Status (Bulk-CSV)
#----------------------------------------------------------------------------------------------------------------
# Autore:				Alan Byrne (18-01-2014)
# URL Originale:		http://gallery.technet.microsoft.com/scriptcenter/Export-a-Licence-b200ca2a
# Versione:				0.2
# Utilizzo:				.\GetMSOLUserLicense.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		18-06-2014
# Modifiche:			
#	0.2- Modificata esclusivamente la formattazione mostrata durante l'esecuzione in PowerShell. Commentata la richiesta di credenziali per la connessione alla console (viene effettuata generalmente prima).
############################################################################################################################

# Define Hashtables for lookup
$Sku = @{
	"DESKLESSPACK" = "Office 365 (Plan K1)"
	"DESKLESSWOFFPACK" = "Office 365 (Plan K2)"
	"LITEPACK" = "Office 365 (Plan P1)"
	"EXCHANGESTANDARD" = "Office 365 Exchange Online Only (Plan P1)"
	"STANDARDPACK" = "Office 365 (Plan E1)"
	"STANDARDWOFFPACK" = "Office 365 (Plan E2)"
	"ENTERPRISEPACK" = "Office 365 (Plan E3)"
	"ENTERPRISEPACKLRG" = "Office 365 (Plan E3)"
	"ENTERPRISEWITHSCAL" = "Office 365 (Plan E4)"
	"STANDARDPACK_STUDENT" = "Office 365 (Plan A1) for Students"
	"STANDARDWOFFPACKPACK_STUDENT" = "Office 365 (Plan A2) for Students"
	"ENTERPRISEPACK_STUDENT" = "Office 365 (Plan A3) for Students"
	"ENTERPRISEWITHSCAL_STUDENT" = "Office 365 (Plan A4) for Students"
	"STANDARDPACK_FACULTY" = "Office 365 (Plan A1) for Faculty"
	"STANDARDWOFFPACKPACK_FACULTY" = "Office 365 (Plan A2) for Faculty"
	"ENTERPRISEPACK_FACULTY" = "Office 365 (Plan A3) for Faculty"
	"ENTERPRISEWITHSCAL_FACULTY" = "Office 365 (Plan A4) for Faculty"
	"ENTERPRISEPACK_B_PILOT" = "Office 365 (Enterprise Preview)"
	"STANDARD_B_PILOT" = "Office 365 (Small Business Preview)"
	}
		
# The Output will be written to this file in the current working directory
$LogFile = "c:\temp\Office_365_Licenses.csv"

# Connect to Microsoft Online
# Import-Module MSOnline
# Connect-MsolService -Credential $Office365credentials

Write-Host "        Office 365: Get User License Status" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -foregroundcolor "red"
Write-Host "          L'operazione potrebbe durare diversi minuti (dipende dal numero di" -foregroundcolor "red"
Write-Host "          utenti registrati all'interno del server Exchange)" -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"

# Get a list of all licences that exist within the tenant
$licensetype = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 1}

# Loop through all licence types found in the tenant
foreach ($license in $licensetype) 
{	
	# Build and write the Header for the CSV file
	$headerstring = "DisplayName;UserPrincipalName;AccountSku"
	
	foreach ($row in $($license.ServiceStatus)) 
	{
		
		# Build header string
		switch -wildcard ($($row.ServicePlan.servicename))
		{
			"EXC*" { $thisLicence = "Exchange Online" }
			"MCO*" { $thisLicence = "Lync Online" }
			"LYN*" { $thisLicence = "Lync Online" }
			"OFF*" { $thisLicence = "Office Profesional Plus" }
			"SHA*" { $thisLicence = "Sharepoint Online" }
			"*WAC*" { $thisLicence = "Office Web Apps" }
			"WAC*" { $thisLicence = "Office Web Apps" }
			default { $thisLicence = $row.ServicePlan.servicename }
		}
		
		$headerstring = ($headerstring + ";" + $thisLicence)
	}
	
	Out-File -FilePath $LogFile -InputObject $headerstring -Encoding UTF8 -append
	
	""
	write-host ("Gathering users with the following subscription: " + $license.accountskuid) -foregroundcolor "yellow"
	""

	# Gather users for this particular AccountSku
	$users = Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses[0].accountskuid.tostring() -eq $license.accountskuid}

	# Loop through all users and write them to the CSV file
	foreach ($user in $users) {
		
		write-host ("     Processing " + $user.displayname)

		$datastring = ($user.displayname + ";" + $user.userprincipalname + ";" + $Sku.Item($user.licenses[0].AccountSku.SkuPartNumber))
		
		foreach ($row in $($user.licenses[0].servicestatus)) {
			
			# Build data string
			$datastring = ($datastring + ";" + $($row.provisioningstatus))
			}
		
		Out-File -FilePath $LogFile -InputObject $datastring -Encoding UTF8 -append
		
	}
}			
""
write-host ("Export Completed. Results available in " + $LogFile) -foregroundcolor "green"
""