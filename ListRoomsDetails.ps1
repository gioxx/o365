<#
OFFICE 365: List Rooms Details
----------------------------------------------------------------------------------------------------------------
Autore:				    GSolone
Versione:			    0.3
Utilizzo:			    .\ListRoomsDetails.ps1
Info:				      https://gioxx.org/tag/o365-powershell
Ultima modifica:	14-11-2022
Modifiche:
  0.3- correggo un parametro nel ciclo di elenco location disponibili e relative sale riunioni per evitare errori di "omonimie" (Get-DistributionGroupMember $roomlist.alias --> Get-DistributionGroupMember $roomlist.PrimarySmtpAddress).
  0.2- da "Rooms Capacity" si passa a "Room Details", lo script integra ora un set di istruzioni per conoscere ogni dettaglio delle sale riunioni, delle liste (location), della capacità e altro ancora.
#>

""; Write-Host "Office 365: List Rooms Details" -f "green";
Write-Host "------------------------------------------"; "";

try	{
	# List Location
	Write-Progress -Activity "Download dati da Exchange" -Status "Elenco le location attualmente gestite..."

	Write-Host "Elenco delle location disponibili:" -f "yellow"
	$Locations = Get-DistributionGroup -RecipientTypeDetails RoomList | ft Name, PrimarySmtpAddress -AutoSize | out-string
	$Locations

	# Mappa corrispondenza Location --> Sala Riunioni (Room List --> Room Mailbox)
	Write-Progress -Activity "Download dati da Exchange" -Status "Mappo la corrispondenza tra location e sale riunioni..."

	Write-Host "Elenco delle location disponibili e relative sale riunioni:" -f "yellow"
	foreach($roomlist in Get-DistributionGroup -RecipientTypeDetails RoomList) {
		$roomlistname = $roomlist.DisplayName
		Get-DistributionGroupMember $roomlist.PrimarySmtpAddress | Select-Object @{n="Room List";e={$roomlistname}},@{n="Room";e={$_.DisplayName}}
	}
	"";""

	# Lista capienza delle sale
	$title = ""
	$message = "Vuoi conoscere la capienza delle sale trovate? (default: NO)"

	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Verifica adesso."
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

	$result = $host.ui.PromptForChoice($title, $message, $options, 1)
	if ($result -eq 0) {
		""; Write-Host "Posti a sedere disponibili nelle sale riunioni rilevate:" -f "yellow";
		Write-Progress -Activity "Download dati da Exchange" -Status "Rilevo la capienza delle sale, attendi..."
		$RoomsCapacity = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "RoomMailbox"} | ft DisplayName,ResourceCapacity -AutoSize | out-string
		$RoomsCapacity
	}
	""
} catch {
	Write-Host "Errore nell'operazione, riprovare." -f "red"
	write-host $error[0]
	return ""
}
