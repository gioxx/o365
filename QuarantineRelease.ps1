<#
	OFFICE 365: Release quarantined messages if the sender is in Exchange Whitelist (single sender or domain)
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.5
	Utilizzo:			.\QuarantineRelease.ps1 ContosoSpamFilterPolicy
						(opzionale, specifica l'indirizzo mittente da cercare in Quarantena) .\QuarantineRelease.ps1 -SenderAddress sender@contoso.com
						(opzionale, specifica il dominio mittente da cercare in Quarantena) .\QuarantineRelease.ps1 -SenderDomain contoso.com
						(opzionale, specifica l'indirizzo mittente da cercare in Quarantena e sbloccalo) .\QuarantineRelease.ps1 -SenderAddress sender@contoso.com -Release
						(opzionale, specifica il dominio mittente da cercare in Quarantena e sbloccalo) .\QuarantineRelease.ps1 -SenderDomain contoso.com -Release
	Info:				https://gioxx.org/tag/o365-powershell
	Fonti utilizzate:	https://social.technet.microsoft.com/wiki/contents/articles/30695.powershell-script-to-identify-quarantine-message-from-specific-domain.aspx
	Ultima modifica:	07-01-2022
	Modifiche:
		0.5- correggo un problema relativo alla query dei messaggi (da mittente specificato) non ancora rilasciati.
		0.4- piccole modifiche ai testi, modifico anche i dati che mostro a video quando faccio preview della coda di Quarantena.
		0.3- inserisco la chiocciola prima del SenderDomain per evitare che vengano sbloccate mail con domini di terzo (o più) livello (capita che si infilino email di phishing / spam).
		0.2rev1- cambio nome script.
		0.2- includo i domini in Whitelist, correggo un problema nel riporto del mittente in sblocco e includo alcuni miglioramenti estetici. Fornisco ora una lista di mail rilasciate al termine dell'intervento.
		0.1- nella modifica dello script in produzione obbligo l'utente a specificare la Spamfilter Policy dalla quale ereditare i mittenti in Whitelist e quando sblocco in maniera automatica i messaggi, lo faccio solo per quelli non ancora rilasciati.
#>

Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $Spamfilter,
	[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $SenderAddress,
	[Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $SenderDomain,
    [switch] $Release
)

if ( ([string]::IsNullOrEmpty($SenderAddress)) -and ([string]::IsNullOrEmpty($SenderDomain)) ) {
	Write-Progress -Activity "Sblocco quarantena da mittenti conosciuti" -Status "Cerco le mail bloccate per TransportRule, attendi ..."
	### Filtro Whitelist Sender --------------------------------------------------------------------
	$SenderWhitelist = Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty AllowedSenders
	foreach ($Sender in $SenderWhitelist) {
		Write-Progress -Activity "Sblocco quarantena da mittenti conosciuti" -Status "Cerco mail dal mittente $($Sender) non ancora rilasciate ..."
		$qm = Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $Sender
		$qmnr = $qm | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ? {$_.QuarantinedUser -ne $null} | ft -AutoSize ReceivedTime,Type,Direction,SenderAddress,Subject,Size,Expires,Released
		$qmnr
		Write-Progress -Activity "Sblocco quarantena da mittenti conosciuti" -Status "Rilascio mail dal mittente $($Sender) ..."
		$qmnr | Release-QuarantineMessage -ReleaseToAll
		$qm | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ft -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
	}
	
	Write-Progress -Activity "Sblocco quarantena da domini conosciuti" -Status "Cerco le mail bloccate per TransportRule, attendi ..."
	### Filtro Whitelist Domains -------------------------------------------------------------------
	$SenderDomainWhitelist = Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty AllowedSenderDomains
	foreach ($Sender in $SenderDomainWhitelist) {
		Write-Progress -Activity "Sblocco quarantena da domini conosciuti" -Status "Cerco mail dal dominio $($Sender) non ancora rilasciate ..."
		$qm = Get-QuarantineMessage -QuarantineTypes TransportRule
		$qmnr = $qm | ? {$_.senderaddress -like "@$Sender"} | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ? {$_.QuarantinedUser -ne $null} | ft -AutoSize ReceivedTime,Type,Direction,SenderAddress,Subject,Size,Expires,Released
		$qmnr
		Write-Progress -Activity "Sblocco quarantena da domini conosciuti" -Status "Rilascio mail dal dominio $($Sender) ..."
		$qmnr | Release-QuarantineMessage -ReleaseToAll
		$qm | ? {$_.senderaddress -like "@$Sender"} | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ft -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
	}
	
} else {
	if ([string]::IsNullOrEmpty($SenderAddress) -eq $false) {
		if ($Release) {
			### Rilascio tutti i messaggi di $SenderAddress in Quarantena
			Write-Progress -Activity "Sblocco quarantena da mittenti conosciuti" -Status "Rilascio mail dal mittente $($SenderAddress) ..."
			Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ? {$_.QuarantinedUser -ne $null} | Release-QuarantineMessage -ReleaseToAll
			Write-Progress -Activity "Sblocco quarantena da mittenti conosciuti" -Status "Verifico mail dal mittente $($SenderAddress) appena rilasciate ..."
			Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ft -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
		} else {
			### Mostro i messaggi di $SenderAddress bloccati in quarantena e non ancora rilasciati
			Write-Progress -Activity "Cerco messaggi in quarantena da mittenti conosciuti" -Status "Elenco mail dal mittente $($SenderAddress) non ancora rilasciate ..."
			Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ft -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
		}
	}
	
	if ([string]::IsNullOrEmpty($SenderDomain) -eq $false) {
		if ($Release) {
			### Rilascio tutti i messaggi di $SenderDomain in Quarantena
			Write-Progress -Activity "Sblocco quarantena da domini conosciuti" -Status "Rilascio mail dal dominio $($SenderDomain) ..."
			Get-QuarantineMessage -QuarantineTypes TransportRule | ? {$_.senderaddress -like "@$SenderDomain"} | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ? {$_.QuarantinedUser -ne $null} | Release-QuarantineMessage -ReleaseToAll
			Write-Progress -Activity "Sblocco quarantena da domini conosciuti" -Status "Verifico mail dal dominio $($SenderDomain) appena rilasciate ..."
			Get-QuarantineMessage -QuarantineTypes TransportRule | ? {$_.senderaddress -like "@$SenderDomain"} | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ft -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
		} else {
			### Mostro i messaggi di $SenderDomain bloccati in quarantena e non ancora rilasciati
			Write-Progress -Activity "Cerco messaggi in quarantena da domini conosciuti" -Status "Elenco mail dal dominio $($SenderDomain) non ancora rilasciate ..."
			$qm = Get-QuarantineMessage -QuarantineTypes TransportRule
			$qm | ? {$_.senderaddress -like "@$SenderDomain"} | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ? {$_.QuarantinedUser -ne $null} | ft -AutoSize ReceivedTime,Type,Direction,SenderAddress,Subject,Size,Expires,Released
		}
	}
}