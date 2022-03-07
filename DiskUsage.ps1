<#
OFFICE 365: Check Mailbox Size on Exchange (Disk Usage)
----------------------------------------------------------------------------------------------------------------
Autore:					          GSolone
Versione:				          0.4
Utilizzo:                 .\DiskUsage.ps1
                          opzionale, modifica posizione export CSV: .\DiskUsage.ps1 C:\temp\Clutter.csv
                          opzionale, singola mailbox specificata: .\DiskUsage.ps1 -Mailbox mario.rossi@contoso.com
                          opzionale, singola mailbox specificata: .\DiskUsage.ps1 -Domain contoso.com
Info:					            https://gioxx.org/tag/o365-powershell
Fonti utilizzate:		      http://www.morgantechspace.com/2015/09/find-office-365-mailbox-size-with-powershell.html
                          https://nchrissos.wordpress.com/2013/06/17/reporting-mailbox-sizes-on-microsoft-exchange-2010/
DEBUG per singolo nodo:   Get-Mailbox -ResultSize 20 | Get-MailboxStatistics | where { $_.OriginatingServer -match "eurprd07.prod.outlook.com"} |
Bug conosciuti:			      se lo script trova una omonimia nel sistema, restituisce errore di tipo: 'La cassetta postale specificata "mario.rossi" non è univoca.'
Ultima modifica:		      02-10-2019
Modifiche:
    0.4- modifica del tipo di esportazione eseguita per permettermi di tirare fuori anche PrimarySmtpAddress e RecipientType oltre all'occupazione e al numero totale di elementi.
    0.3- aggiungo delimitatore ";" all'export-CSV.
    0.2- aggiungo possibilità di esportare le statistiche solo delle caselle appartenenti a uno specifico dominio.
    0.1 rev2- migliorata formattazione estrazione dati in CSV.
    0.1 rev1- correzioni minori.
#>

#Verifica parametri da prompt
Param(
  [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)][string] $CSV,
  [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)][string] $Mailbox,
  [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)][string] $Domain
)

$DataOggi = Get-Date -format yyyyMMdd

""; Write-Host "        Office 365: Check Mailbox Size on Exchange (Disk Usage)" -f "green"
Write-Host "        ------------------------------------------"

# Mailbox non specificata, estrazione dati completa
if ([string]::IsNullOrEmpty($Mailbox)) {
  # Dominio non specificato, estrazione dati completa
  if ([string]::IsNullOrEmpty($Domain)) {
    <#
    Puoi modificare il valore $CSV per impostare un diverso nome del file CSV che verrà
    esportato dallo script (solo ciò che c'è tra le virgolette, ad esempio
    $CSV = "C:\temp\CSV.csv" (per modificare anche la cartella di esportazione),
    oppure $CSV = "Permessi.csv" per salvare il file nella stessa cartella dello script.
    ATTENZIONE: utilizza (per comodità) nomi diversi nel caso in cui lo script esporti i permessi
    delle caselle ShaRed piuttosto che quelle personali.
    #>
    if ([string]::IsNullOrEmpty($CSV)) {
      $CSV = "C:\temp\DiskUsage_$DataOggi.csv"
    }

    Write-Host "          ATTENZIONE:" -f "Red"
    Write-Host "          L'operazione può richiedere MOLTO tempo, dipende dal numero di utenti"
    Write-Host "          da verificare e modificare all'interno della Directory, porta pazienza!"
    ""
    Write-Host "          * Per modificare la posizione del file CSV esportato, rilancia lo script con parametro"
    Write-Host "            -CSV C:\export.csv" -nonewline -f "Yellow"; Write-Host " (es. " -nonewline; Write-Host ".\DiskUsage.ps1 -CSV C:\export.csv" -nonewline -f "Yellow"; Write-Host ")"
    Write-Host "          * Per analizzare una singola mailbox, rilancia lo script con parametro"
    Write-Host "            -Mailbox mario.rossi@contoso.com" -nonewline -f "Yellow"; Write-Host " (es. " -nonewline; Write-Host ".\DiskUsage.ps1 -Mailbox mario.rossi@contoso.com" -nonewline -f "Yellow"; Write-Host ")"
    Write-Host "          * Per analizzare uno specifico dominio, rilancia lo script con parametro"
    Write-Host "            -Domain contoso.com" -nonewline -f "Yellow"; Write-Host " (es. " -nonewline; Write-Host ".\DiskUsage.ps1 -Domain contoso.com" -nonewline -f "Yellow"; Write-Host ")"
    Write-Host "        ------------------------------------------"
    ""

    Function Pause($M="        Premi un tasto continuare (CTRL+C per annullare)") {
      If($psISE) {
        $S=New-Object -ComObject "WScript.Shell"
        $B=$S.Popup("Fai clic su OK per continuare.",0,"In attesa dell'amministratore",0)
        Return
      }
      Write-Host -NoNewline $M;
      $I=16,17,18,20,91,92,93,144,145,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183;
      While($K.VirtualKeyCode -Eq $Null -Or $I -Contains $K.VirtualKeyCode) {
        $K=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
      }
    }
    Pause

    try {
      ""; ""; Write-Host "        File CSV di destinazione: " -nonewline; Write-Host $CSV -f "Yellow"; ""
      Write-Host "        A long time left, grab a Snickers!" -f "Yellow"
      Write-Progress -Activity "Download dati da Exchange" -Status "Scarico occupazione disco delle caselle registrate nel sistema..."
      # Analisi occupazione Mailbox, sort, salvataggio su file

      Get-Mailbox -ResultSize Unlimited |
      Select-Object DisplayName,
      PrimarySmtpAddress,RecipientTypeDetails,
      @{Name='TotalItemSize(GB)'; expression={[math]::Round((((Get-MailboxStatistics $_.PrimarySmtpAddress).TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},
      @{Name='ItemCount'; expression={(Get-MailboxStatistics $_.PrimarySmtpAddress).ItemCount}} |
      Export-Csv $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"

      ""; Write-Host "Ho terminato l'esportazione dei dati." -f "Green"

      # Chiedo se visualizzare il file CSV Generato
      $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Apri ora il file CSV"
      $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Termina lo script senza aprire il file CSV"
      $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
      $OpenCSV = $host.ui.PromptForChoice("Devo aprire il file CSV generato?", $message, $options, 1)
      if ($OpenCSV -eq 0) { Invoke-Item $CSV }
      ""
    }
    catch
    {
      Write-Host "Errore nell'operazione, riprovare." -f "Red"
      write-host $error[0]
      return ""
    }
  } else {
    # Se non è specificato il nome del CSV, lo genero con la data odierna e il dominio richiesto
    if ([string]::IsNullOrEmpty($CSV)) {
      $CSV = "C:\temp\DiskUsage_$Domain_$DataOggi.csv"
    }

    ""; ""; Write-Host "        File CSV di destinazione: " -nonewline; Write-Host $CSV -f "Yellow";
    Write-Host "        Dominio da analizzare specificato: " -nonewline; Write-Host *$($Domain) -f "Yellow"

    Get-Mailbox -ResultSize Unlimited | where {$_.PrimarySmtpAddress -like "*" + $Domain} |
    Select-Object DisplayName,
    PrimarySmtpAddress,RecipientTypeDetails,
    @{Name='TotalItemSize(GB)'; expression={[math]::Round((((Get-MailboxStatistics $_.PrimarySmtpAddress).TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},
    @{Name='ItemCount'; expression={(Get-MailboxStatistics $_.PrimarySmtpAddress).ItemCount}} |
    Export-Csv $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    ""; Write-Host "Ho terminato l'esportazione dei dati." -f "Green"

    # Chiedo se visualizzare il file CSV Generato
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Apri ora il file CSV"
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Termina lo script senza aprire il file CSV"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
    $OpenCSV = $host.ui.PromptForChoice("Devo aprire il file CSV generato?", $message, $options, 1)
    if ($OpenCSV -eq 0) { Invoke-Item $CSV }
    ""
  }
} else {
  # Mailbox specificata, estrazione dati singola
  ""; Write-Host "        Mailbox da analizzare specificata: " -nonewline; Write-Host $Mailbox -f "Yellow"
  Get-Mailbox $Mailbox | Get-MailboxStatistics |
  Select-Object -Property @{label="User";expression={$_.DisplayName}},
  @{label="Total Messages";expression= {$_.ItemCount}},
  @{label="Total Size (GB)";expression={[math]::Round(`
    # Trasformo in GB
    ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}}
  }
