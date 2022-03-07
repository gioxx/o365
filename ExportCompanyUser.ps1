<#
OFFICE 365: Export Company Users
----------------------------------------------------------------------------------------------------------------
Autore:				    GSolone
Versione:			    0.9
Utilizzo:			    .\ExportCompanyUsers.ps1
                  (opzionale, passaggio dati da prompt) .\ExportCompanyUsers.ps1 -RicercaDominio contoso.com
                  (opzionale, passaggio dati da prompt) .\ExportCompanyUsers.ps1 -RicercaCompany "Contoso S.r.l." (comprensivo di virgolette)
Info:				      https://gioxx.org/tag/o365-powershell
Ultima modifica:	06-06-2019
Modifiche:
    0.9- aggiungo delimitatore ";" all'export-CSV.
    0.8- riscritto, snellito e velocizzato.
    0.7 rev1- ho solo corretto l'Out-String di uscita dei dati a video per RicercaDominio e RicercaCompany (ora -AutoSize)
    0.7- ho modificato la modalità di esportazione dati, ora i risultati vanno in un file CSV. Ho commentato le righe relative all'aggiunta di data e ora dell'esportazione.
    0.6- prevedo utilizzo del parametro -RicercaCompany per filtrare un campo basato sulla Company e non sul Mail Domain.
    0.5- lo script accetta adesso i parametri passati da riga di comando (-RicercaDominio). Nuovo metodo di output dei dati trovati, ricerco prima le caselle, poi per ciascuna casella ricavo i dati che mi servono direttamente dallo User, permettendo così l'esportazione anche del campo Company. Chiedo se esportare i risultati in CSV (al posto di farlo per default).
    0.4- eliminato calcolo avanzamento download dati.
    0.3- correzione minore: la ricerca viene effettuata sullo specifico dominio in ingresso, un eventuale sottodominio deve essere dichiarato (esempio: contoso.com nella ricerca non mostrerà i risultati di dep1.contoso.com nell'eventualità dep1 fosse un suo sottodominio).
    0.2- stato di avanzamento in lettura / scrittura dati). Modificato il $_.EmailAddresses in $_.PrimarySmtpAddress per mettere la Company in base all'indirizzo di posta principale e non considerare eventuali alias
#>

#Verifica parametri da prompt
Param(
  [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)][string] $RicercaDominio,
  [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)][string] $RicercaCompany
)

#Main
Function Main {
  ""; Write-Host "        Office 365: Export Company Users" -f "Green"
  Write-Host "        ------------------------------------------"

  if ( [string]::IsNullOrEmpty($RicercaCompany) -and [string]::IsNullOrEmpty($RicercaDominio) )
  {
    # Manca dominio di ricerca e non è specificato la Company, chiedo a video il dominio da analizzare
    $RicercaDominio = Read-Host "        Dominio da analizzare (esempio: contoso.com) "
  }

  #Ricerca basata sul campo Company
  if ( [string]::IsNullOrEmpty($RicercaCompany) -eq $False )
  {
    Write-Host "        Azienda di ricerca: " -f "White" -nonewline; Write-Host "$RicercaCompany" -f "Green"
    Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle con appartenenti al gruppo che mi hai richiesto, attendi..."

    ""; Get-Recipient -ResultSize Unlimited | where {$_.Company -eq "$RicercaCompany"} | Select-Object DisplayName, PrimarySmtpAddress, Company, City | Sort-Object PrimarySmtpAddress
    ""

    # Chiedo se esportare i risultati in un file CSV
    $ExportList = "C:\temp\$RicercaCompany.csv"
    $message = "Devo esportare i risultati in $ExportList ?"
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Esporta ora l'elenco"
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non desidero esportare ora l'elenco"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
    $ExportCSV = $host.ui.PromptForChoice("", $message, $options, 1)
    if ($ExportCSV -eq 0) {
      Write-Host "Esporto l'elenco in C:\temp\$RicercaCompany.csv e apro il file (salvo errori)." -f "Green"
      $Today = [string]::Format( "{0:dd/MM/yyyy}", [datetime]::Now.Date )
      $TimeIs = (get-date).tostring('HH:mm:ss')
      Get-Recipient -ResultSize Unlimited | where {$_.Company -eq "$RicercaCompany"} | Select-Object DisplayName, PrimarySmtpAddress, Company, City | Export-CSV $ExportList -force -notypeinformation -Delimiter ";"
      $a = Get-Content $ExportList
      $b = "Esportazione utenti $RicercaCompany - $Today alle ore $TimeIs"
      Invoke-Item $ExportList
    }
    ""; Write-Host "Done." -f "Green"; "";
  }

  # Ricerca basata sul dominio
  if ( [string]::IsNullOrEmpty($RicercaDominio) -eq $False )
  {
    Write-Host "        Dominio di ricerca: " -f "White" -nonewline; Write-Host "$RicercaDominio" -f "Green"
    Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle con il dominio che mi hai richiesto, attendi..."

    ""; Get-Recipient -ResultSize Unlimited | where {$_.Company -eq "$RicercaCompany"} | Select-Object DisplayName, PrimarySmtpAddress, Company, City | Sort-Object PrimarySmtpAddress
    ""

    # Chiedo se esportare i risultati in un file CSV
    $ExportList = "C:\temp\$RicercaDominio.csv"
    $message = "Devo esportare i risultati in $ExportList ?"
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
    $ExportCSV = $host.ui.PromptForChoice("", $message, $options, 1)
    if ($ExportCSV -eq 0) {
      Write-Host "Esporto l'elenco in C:\temp\$RicercaDominio.csv e apro il file (salvo errori)." -f "Green"
      $Today = [string]::Format( "{0:dd/MM/yyyy}", [datetime]::Now.Date )
      $TimeIs = (get-date).tostring('HH:mm:ss')
      Get-Recipient -ResultSize Unlimited | where {$_.Company -eq "$RicercaCompany"} | Select-Object DisplayName, PrimarySmtpAddress, Company, City | Export-CSV $ExportList -force -notypeinformation -Delimiter ";"
      $a = Get-Content $ExportList
      $b = "Esportazione utenti $RicercaDominio - $Today alle ore $TimeIs"
      Invoke-Item $ExportList
    }
    ""; Write-Host "Done." -f "Green"; "";
  }
}

# Start script
. Main
