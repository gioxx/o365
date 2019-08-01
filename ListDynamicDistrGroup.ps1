<#
	OFFICE 365: Show Dynamic Distribution Group Users
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.5
	Utilizzo:			.\ListDynamicDistrGroup.ps1
						(opzionale, passaggio dati da prompt) .\ListDynamicDistrGroup.ps1 group@contoso.com
						(opzionale, passaggio dati da prompt) .\ListDynamicDistrGroup.ps1 group@contoso.com -CSV C:\temp\group.CSV
	Info:				https://gioxx.org/tag/o365-powershell
	Ultima modifica:	06-06-2019
	Modifiche:
		0.5- aggiungo delimitatore ";" all'export-CSV.
		0.4- Select FirstName, LastName, DisplayName, Name al posto del vecchio Select Name. Metto a posto alcuni particolari "estetici" e snellisco le istruzioni utilizzate.
		0.3 rev2- Sort-Object sull'output. Esportazione in CSV e non più su file di testo.
		0.3 rev1- mostro a video le informazioni che prima esportavo esclusivamente nel CSV finale.
		0.3- l'esportazione su CSV include anche l'indirizzo SMTP primario.
		0.2- lo script accetta adesso i parametri passati da riga di comando (-RicercaGruppo e -CSV), e permette di esportare il risultato della query su file CSV.
#>

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $RicercaGruppo, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $CSV
)

#Main
Function Main {
	
	""; Write-Host "        Office 365: Show Dynamic Distribution Group Users" -f "green"
	Write-Host "        -----------------------------------------------------------------------------------------"

	if ( [string]::IsNullOrEmpty($RicercaGruppo) ) {
		""
		#Mancano i dettagli da prompt, li richiedo a video
		$RicercaGruppo = Read-Host "        Mostra utenti del gruppo (esempio: Contoso srl - Utenti)"
	}

	""; Write-Host "        Gruppo da analizzare: " -nonewline -f "yellow"; Write-Host "$RicercaGruppo "
	
	if ( [string]::IsNullOrEmpty($CSV) -eq 0 ) {
		Write-Host "        File CSV di destinazione: " -nonewline -f "yellow"; Write-Host "$CSV "
	}
	
	try {
		Write-Progress -Activity "Download dati da Exchange" -Status "Esporto utenti in $RicercaGruppo, attendi..."
		$members = Get-DynamicDistributionGroup -Identity $RicercaGruppo
		""; Write-Host "        Anteprima dei risultati (indirizzo di posta elettronica):" -f "Yellow"; "";
		Get-Recipient -RecipientPreviewFilter $members.RecipientFilter | ft PrimarySmtpAddress
		
		if ( [string]::IsNullOrEmpty($CSV) -eq 0 ) {
			# Esportazione risultati in CSV
			Get-Recipient -RecipientPreviewFilter $members.RecipientFilter | Select FirstName, LastName, DisplayName, Name, PrimarySmtpAddress, Company, City | Export-CSV $CSV -notypeinformation -force -Delimiter ";"
			# Richiedo apertura file CSV
			$message = "Devo aprire il file $($CSV)? (default: NO)"
			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Si."
			$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
			$result = $host.ui.PromptForChoice("", $message, $options, 1)
			if ($result -eq 0) { Invoke-Item $CSV }
		}
	} catch {
		Write-Host "Errore nell'operazione, riprovare." -f "red"
		Write-Host $error[0]
		return ""
	}
}

# Start script
. Main