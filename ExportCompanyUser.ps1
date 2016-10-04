############################################################################################################################
# OFFICE 365: Export Company Users
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.6
# Utilizzo:				.\ExportCompanyUsers.ps1
#						(opzionale, passaggio dati da prompt) .\ExportCompanyUsers.ps1 contoso.com
#						(opzionale, passaggio dati da prompt) .\ExportCompanyUsers.ps1 -RicercaCompany "Contoso S.r.l."
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		09-05-2016
# Modifiche:			
#	0.6- prevedo utilizzo del parametro -RicercaCompany per filtrare un campo basato sulla Company e non sul Mail Domain.
#	0.5- lo script accetta adesso i parametri passati da riga di comando (-RicercaDominio). Nuovo metodo di output dei dati trovati, ricerco prima le caselle, poi per ciascuna casella ricavo i dati che mi servono direttamente dallo User, permettendo così l'esportazione anche del campo Company. Chiedo se esportare i risultati in CSV (al posto di farlo per default).
#	0.4- Eliminato calcolo avanzamento download dati.
#	0.3- Correzione minore: la ricerca viene effettuata sullo specifico dominio in ingresso, un eventuale sottodominio deve essere dichiarato (esempio: contoso.com nella ricerca non mostrerà i risultati di dep1.contoso.com nell'eventualità dep1 fosse un suo sottodominio).
#	0.2- Stato di avanzamento in lettura / scrittura dati). Modificato il $_.EmailAddresses in $_.PrimarySmtpAddress per mettere la Company in base all'indirizzo di posta principale e non considerare eventuali alias
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $RicercaDominio,
	[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $RicercaCompany
)

#Main
Function Main {

	""
	Write-Host "        Office 365: Export Company Users" -f "Green"
	Write-Host "        ------------------------------------------"
	
	if ( [string]::IsNullOrEmpty($RicercaCompany) -and [string]::IsNullOrEmpty($RicercaDominio) )
	{
		# Manca dominio di ricerca e non è specificato la Company, chiedo a video il dominio da analizzare
		$RicercaDominio = Read-Host "        Dominio da analizzare (esempio: domain.tld) "
	}
	
	#Ricerca basata sul campo Company
	if ( [string]::IsNullOrEmpty($RicercaCompany) -eq $False )
		{
			Write-Host "        Azienda di ricerca: " -f "White" -nonewline; Write-Host "$RicercaCompany" -f "Green"
			Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle con appartenenti al gruppo che mi hai richiesto, attendi..."
			
			$RicercaMailbox = Get-User -ResultSize Unlimited | where {$_.Company -eq "$RicercaCompany"}
			$RicercaMailbox | foreach {Get-User $_.Name} | ft DisplayName, WindowsEmailAddress, Company, City | Out-String
			
			# Chiedo se esportare i risultati in un file CSV
			$ExportList = "C:\temp\$RicercaCompany.txt"
			$message = "Devo esportare i risultati in $ExportList ?"
			$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Esporta ora l'elenco"
			$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non desidero esportare ora l'elenco"
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
			$ExportCSV = $host.ui.PromptForChoice("", $message, $options, 1)
			if ($ExportCSV -eq 0) { 
				Write-Host "Esporto l'elenco in C:\temp\$RicercaCompany.txt e apro il file (salvo errori)." -f "Green"
				$Today = [string]::Format( "{0:dd/MM/yyyy}", [datetime]::Now.Date )
				$TimeIs = (get-date).tostring('HH:mm:ss')		
				$RicercaMailbox | foreach {Get-User $_.Name} | ft DisplayName, WindowsEmailAddress, Company, City > $ExportList
				$a = Get-Content $ExportList
				$b = "Esportazione utenti $RicercaCompany - $Today alle ore $TimeIs"
				#Set-Content $ExportList –value $b, $a[0..18]
				Set-Content $ExportList –value $b, $a
				Invoke-Item $ExportList
			}
			""		
			Write-Host "Done." -f "Green"
			""
		}
	
	# Ricerca basata sul dominio
	if ( [string]::IsNullOrEmpty($RicercaDominio) -eq $False )
		{
			Write-Host "        Dominio di ricerca: " -f "White" -nonewline; Write-Host "$RicercaDominio" -f "Green"
			Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle con il dominio che mi hai richiesto, attendi..."
			
			#$RicercaMailbox= Get-Mailbox -ResultSize unlimited | where {$_.EmailAddresses -like "*@" + $RicercaDominio}
			$RicercaMailbox = Get-Mailbox -ResultSize Unlimited | where {$_.PrimarySmtpAddress -like "*@" + $RicercaDominio}
			#$RicercaMailbox | FT DisplayName, UserPrincipalName, Company
			$RicercaMailbox | foreach {Get-User $_.Name} | ft DisplayName, WindowsEmailAddress, Company, City | Out-String
			
			# Chiedo se esportare i risultati in un file CSV
			$ExportList = "C:\temp\$RicercaDominio.txt"
			$message = "Devo esportare i risultati in $ExportList ?"
			$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
			$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
			$ExportCSV = $host.ui.PromptForChoice("", $message, $options, 1)
			if ($ExportCSV -eq 0) { 
				Write-Host "Esporto l'elenco in C:\temp\$RicercaDominio.txt e apro il file (salvo errori)." -f "Green"
				$Today = [string]::Format( "{0:dd/MM/yyyy}", [datetime]::Now.Date )
				$TimeIs = (get-date).tostring('HH:mm:ss')		
				$RicercaMailbox | foreach {Get-User $_.Name} | ft DisplayName, WindowsEmailAddress, Company, City > $ExportList
				$a = Get-Content $ExportList
				$b = "Esportazione utenti $RicercaDominio - $Today alle ore $TimeIs"
				#Set-Content $ExportList –value $b, $a[0..18]
				Set-Content $ExportList –value $b, $a
				Invoke-Item $ExportList
			}
			""		
			Write-Host "Done." -f "Green"
			""
		}
}

# Start script
. Main