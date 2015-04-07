############################################################################################################################
# OFFICE 365: List Rooms Capacity
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\ListRoomsCapacity.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		05-02-2015
# Modifiche:			-
############################################################################################################################

""
Write-Host "        Office 365: List Rooms Capacity" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
""

try
	{
		Write-Host "Ricerco tutte le sale riunioni e ne elenco la capacità, attendi." -foregroundcolor "yellow"
		""
		Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "RoomMailbox"} | ft DisplayName,ResourceCapacity -AutoSize
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
		write-host $error[0]
		return ""
	}