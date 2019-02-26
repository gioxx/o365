<#
	OFFICE 365: Automatic release quarantined messages if the sender is in Exchange Whitelist
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.1
	Utilizzo:			.\AutomaticQuarantineRelease.ps1 ContosoSpamFilterPolicy
						(opzionale, specifica il mittente da cercare in Quarantena) .\AutomaticQuarantineRelease.ps1 -SenderAddress sender@contoso.com
						(opzionale, specifica il mittente da cercare in Quarantena e sbloccalo) .\AutomaticQuarantineRelease.ps1 -SenderAddress sender@contoso.com -Release
	Info:				https://gioxx.org/tag/o365-powershell
	Ultima modifica:	25-02-2019
	Modifiche:
		0.1- nella modifica dello script in produzione obbligo l'utente a specificare la Spamfilter Policy dalla quale ereditare i mittenti in Whitelist e quando sblocco in maniera automatica i messaggi, lo faccio solo per quelli non ancora rilasciati.
#>

Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $Spamfilter,
	[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $SenderAddress,
    [switch] $Release
)

if ( [string]::IsNullOrEmpty($SenderAddress) ) {
	Write-Progress -Activity "Sblocco quarantena da mittenti conosciuti" -Status "Cerco le mail bloccate per TransportRule, attendi ..."
	### Filtro Whitelist Sender --------------------------------------------------------------------
	$SenderWhitelist = Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty AllowedSenders
	foreach ($Sender in $SenderWhitelist) {
		Write-Progress -Activity "Sblocco quarantena da mittenti conosciuti" -Status "Cerco mail dal mittente $($Sender) non ancora rilasciati ..."
		$qm = Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $Sender
		$qmnr = $qm | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ? {$_.QuarantinedUser -ne $null}
		Write-Progress -Activity "Sblocco quarantena da mittenti conosciuti" -Status "Rilascio mail dal mittente $($d) ..."
		$qmnr | Release-QuarantineMessage -ReleaseToAll
	}
} else {
	if ($Release) {
		### Rilascio tutti i messaggi di $SenderAddress in Quarantena (Warning a video se già rilasciati)
		Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | Release-QuarantineMessage -ReleaseToAll
	} else {
		### Mostro i messaggi di $SenderAddress bloccati in quarantena e non ancora rilasciati
		$qm = Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress
		$qm | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ? {$_.QuarantinedUser -ne $null}
	}
}