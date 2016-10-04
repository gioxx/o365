############################################################################################################################
# OFFICE 365: Show Dynamic Distribution Group Users
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.3 rev1
# Utilizzo:				.\ListDynamicDistrGroup.ps1
#						(opzionale, passaggio dati da prompt) .\ListDynamicDistrGroup.ps1 group@contoso.com
#						(opzionale, passaggio dati da prompt) .\ListDynamicDistrGroup.ps1 group@contoso.com -TXT C:\temp\group.txt
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		16-09-2016
# Modifiche:			
#	0.3 rev1- mostro a video le informazioni che prima esportavo esclusivamente nel TXT finale.
#	0.3- l'esportazione su TXT include anche l'indirizzo SMTP primario.
#	0.2- lo script accetta adesso i parametri passati da riga di comando (-RicercaGruppo e -TXT), e permette di esportare il risultato della query su file TXT.
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $RicercaGruppo, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $TXT
)

#Main
Function Main {
	
	""
	Write-Host "        Office 365: Show Dynamic Distribution Group Users" -f "green"
	Write-Host "        -----------------------------------------------------------------------------------------"

	if ( [string]::IsNullOrEmpty($RicercaGruppo) )
	{
		""
		#MANCANO I DETTAGLI DA PROMPT, LI RICHIEDO A VIDEO
		$RicercaGruppo = Read-Host "        Mostra utenti del gruppo (esempio: Contoso srl - Utenti)"
	}

	""
	Write-Host "        Gruppo da analizzare: " -nonewline -f "yellow"; Write-Host "$RicercaGruppo "
	
	try
	{
		Write-Progress -Activity "Download dati da Exchange" -Status "Esporto utenti in $RicercaGruppo, attendi..."
		$members = Get-DynamicDistributionGroup -Identity $RicercaGruppo
		Get-Recipient -RecipientPreviewFilter $members.RecipientFilter | ft Name, PrimarySmtpAddress, Company, City
		""
		
		# Esportazione risultati in TXT
		$message = "Vuoi esportare il risultato in un file TXT? (default: NO)"
		$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Si."
		$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
		$result = $host.ui.PromptForChoice("", $message, $options, 1)
		if ($result -eq 0) { 
			""
			if ( [string]::IsNullOrEmpty($TXT) )
				{
					# Directory esportazione TXT non specificata, utilizzo default
					$TXT = "C:\temp\$RicercaGruppo.txt"
				}
			Write-Host "Esporto i risultati in $TXT :" -nonewline -f "yellow"
			Get-Recipient -RecipientPreviewFilter $members.RecipientFilter | ft Name, PrimarySmtpAddress, Company, City > $TXT
			Write-Host " fatto" -f "green"
			""
			# Richiedo apertura file TXT
			$message = "Devo aprire il file TXT generato? (default: NO)"
			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Si."
			$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non ora."
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
			$result = $host.ui.PromptForChoice("", $message, $options, 1)
			if ($result -eq 0) { Invoke-Item $TXT }
		}
		""
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -f "red"
		write-host $error[0]
		return ""
	}
	
}

# Start script
. Main