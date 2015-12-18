############################################################################################################################
# OFFICE 365: Explode all Groups
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\ExplodeAllGroups-CSV.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		02-10-2015
# Fonti utilizzate:		https://gallery.technet.microsoft.com/office/List-all-Users-Distribution-7f2013b2
#						https://gallery.technet.microsoft.com/office/Export-all-distribution-707c27eb
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	################################################################################################
	# Puoi modificare questi valori per impostare un diverso nome dei file CSV che verranno
	# esportati dallo script (solo ciò che c'è tra le virgolette, ad esempio
	# $ExportDistrGroup = "C:\temp\DistributionGroupMembers.csv" (per modificare anche la cartella di esportazione), 
	# oppure $ExportDistrGroup = "DistributionGroupMembers.csv" per salvare il file nella stessa cartella dello script.
	$ExportDistrGroup = "C:\temp\DistributionGroupMembers.csv"
	$ExportDynOutputFile = "C:\temp\DynDistributionGroupMembers.csv"
	################################################################################################

	""
	Write-Host "        Office 365: Explode all Groups" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "        Lo script elenca tutti i gruppi (dinamici e non) presenti su Exchange " -f "white"
	Write-Host "        elecandone inoltre tutti i membri rilevati al loro interno.     " -f "white"
	Write-Host "        Il risultato verrà poi esportato su due file CSV in C:\Temp     " -f "white"
	Write-Host "        (posizione e nome dei file possono variare modificando questo script)" -f "white"
	""
	Write-Host "        Posizione dei file rilevata:"
	Write-Host "		  - $ExportDistrGroup" -f "green"
	Write-Host "		  - $ExportDynOutputFile" -f "green"
	""
	Write-Host "        Premi un tasto qualsiasi per continuare..."
	[void][System.Console]::ReadKey($true)
	
	try
	{
		""
		
		#---------------------------------------------------------------------------------------
		# Gruppi di Distribuzione standard
		#---------------------------------------------------------------------------------------
		
		Write-Progress -Activity "Download dati da Exchange" -Status "Esporto gruppi e utenti in $ExportDistrGroup, attendi..."

		#Preparazione file di output
		Out-File -FilePath $ExportDistrGroup -InputObject "Distribution Group DisplayName,Distribution Group Email,Member DisplayName, Member Email, Member Type" -Encoding UTF8

		#Ricerca dei gruppi e verifica utenti al loro interno
		$objDistributionGroups = Get-DistributionGroup -ResultSize Unlimited  
		Foreach ($objDistributionGroup in $objDistributionGroups) 
		{
			
			""
			Write-Host "Verifico $($objDistributionGroup.DisplayName)..." -f "Yellow"
			$objDGMembers = Get-DistributionGroupMember -Identity $($objDistributionGroup.PrimarySmtpAddress)
			""
			Write-Host "Rilevati $($objDGMembers.Count) utenti nel gruppo..." -f "Green"
			
			#Verifica caratteristiche utente e risultato a video / CSV
			Foreach ($objMember in $objDGMembers)
			{
				Out-File -FilePath $ExportDistrGroup -InputObject "$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" -Encoding UTF8 -append
				Write-Host "`t$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)"
			}
		}
		
		#---------------------------------------------------------------------------------------
		# Gruppi di Distribuzione Dinamici
		#---------------------------------------------------------------------------------------

		Write-Progress -Activity "Download dati da Exchange" -Status "Esporto gruppi dinamici e utenti in $ExportDynOutputFile, attendi..."
		
		#Preparazione file di output
		Out-File -FilePath $ExportDynOutputFile -InputObject "Dynamic Distribution Group DisplayName,Dynamic Distribution Group Email,Member DisplayName, Member Email" -Encoding UTF8
		
		#Ricerca dei gruppi e verifica utenti al loro interno
		$objDynDistributionGroups = Get-DynamicDistributionGroup -ResultSize Unlimited
		Foreach ($objDynDistributionGroup in $objDynDistributionGroups)
		{
			""
			Write-Host "Verifico $($objDynDistributionGroup.DisplayName)..." -f "Yellow"
			$objDGMembers = Get-Recipient -RecipientPreviewFilter $objDynDistributionGroup.RecipientFilter -ResultSize Unlimited
			""
			Write-Host "Rilevati $($objDGMembers.Count) utenti nel gruppo..." -f "Green"
			
			#Verifica caratteristiche utente e risultato a video / CSV
			Foreach ($objMember in $objDGMembers)
			{
				Out-File -FilePath $ExportDynOutputFile -InputObject "$($objDynDistributionGroup.DisplayName),$($objDynDistributionGroup.PrimarySmtpAddress),$($objMember.Name),$($objMember.PrimarySMTPAddress)" -Encoding UTF8 -append
				Write-Host "`t$($objDynDistributionGroup.DisplayName),$($objDynDistributionGroup.PrimarySMTPAddress),$($objMember.Name),$($objMember.PrimarySMTPAddress)"
			}
		}
		
		#Apro i file di Output con l'editor predefinito (es. Excel)
		Invoke-Item $ExportDistrGroup
		Invoke-Item $ExportDynOutputFile
		
		""
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