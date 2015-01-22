############################################################################################################################
# OFFICE 365: Show Dynamic Distribution Group Users
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\ListDynamicDistrGroup.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		22-01-2015
# Modifiche:			-
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Show Dynamic Distribution Group Users" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	$RicercaGruppo = Read-Host "Mostra utenti del gruppo (esempio: Emmelibri srl - Utenti) "
	
	try
	{
		""
		Write-Host "Ricerco le caselle di posta facenti parte del gruppo indicato, attendi." -foregroundcolor "yellow"
		$members = Get-DynamicDistributionGroup -Identity $RicercaGruppo
		Get-Recipient -RecipientPreviewFilter $members.RecipientFilter
		
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