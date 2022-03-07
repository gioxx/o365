############################################################################################################################
# OFFICE 365: Set New Primary SMTP Address (CSV)
#----------------------------------------------------------------------------------------------------------------
# Autore:				       GSolone
# Versione:				     0.2 rev1
# Utilizzo:				     .\SetNewPrimarySMTPAddress-CSV.ps1
# Info:					      https://gioxx.org/tag/o365-powershell
# Ultima modifica:		10-03-2016
# Modifiche:
#   0.2 rev1 - aggiunta funzione di Pausa per evitare di intercettare il tasto CTRL.
#   0.2- lo script accetta ora il parametro CSV da riga di comando (esempio: .\SetNewPrimarySMTPAddress-CSV.ps1 C:\temp\Utenti.csv). Aggiunto il blocco di modifica MsolUserPrincipalName oltre l'indirizzo principale SMTP della casella di posta.
############################################################################################################################

#Verifica parametri da prompt
Param(
  [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]
  [string] $CSV
)

#Main
Function Main {
  ""; Write-Host "        Office 365: Add New PrimarySMTPAddress" -f "green"
  Write-Host "        ------------------------------------------"
  Write-Host "         Costruire il file CSV con in colonna 1 l'attuale indirizzo di posta elettronica" -f "white"
  Write-Host "         e in colonna 2 il nuovo indirizzo da far diventare primario." -f "white"
  Write-Host "         Il titolo della prima colonna dovrà essere " -nonewline; Write-Host "indirizzo_attuale" -f "yellow" -nonewline; Write-Host ", la seconda " -f "white" -nonewline; Write-Host "nuovo_indirizzo" -f "yellow"
  Write-Host "         Il file dovrà essere salvato come Utenti.csv e trovarsi nella stessa posizione di questo script" -f "white"; ""
  Write-Host "         In caso contrario sarà necessario lanciare lo script con parametro posizione del file CSV" -f "white"
  Write-Host "         esempio: .\SetNewPrimarySMTPAddress-CSV.ps1 C:\temp\Utenti.csv" -f "white"; ""
  Write-Host "         CSV di esempio:" -f "white"
  Write-Host "         indirizzo_attuale,nuovo_indirizzo" -f "gray"
  Write-Host "         test_1@contoso.onmicrosoft.com,test_1@contoso.com" -f "gray"
  Write-Host "         test_2@contoso.onmicrosoft.com,test_2@contoso.com" -f "gray"; ""

  Function Pause($M="Premi un tasto continuare (CTRL+C per annullare)") {
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

  try
  {
    Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle elencate nel file CSV, attendi..."

    #Nessun parametro passato, cerco utenti.csv nella stessa cartella dello script
    if ([string]::IsNullOrEmpty($CSV) -eq $true) { $CSV = ".\Utenti.csv" }

    #Procedo con il lavoro di modifica indirizzo primario
    import-csv $CSV | ForEach-Object {
      $user = Get-Mailbox -Identity $_.indirizzo_attuale

      <#
        DEBUG
        ""; Write-Host "Se vedi questo testo è attivo il blocco debug, entra nello script e commentalo se necessario. Verifica user letto:" -f "red"
        $user
      #>

      Write-Progress -activity "Modifica Primary SMTP Address" -status "Modifico $_.indirizzo_attuale"

      #Recupero dati utente
      $OldPrimarySMTPAddress = $_.indirizzo_attuale
      $NewPrimarySMTPAddress = $_.nuovo_indirizzo

      #Modifica impostazioni casella di posta (indirizzo primario)
      ""; ""; Write-Host "		" -nonewline; Write-Host "Applicato nuovo indirizzo: $NewPrimarySMTPAddress" -b "green" -f "black"; ""
      $user.EmailAddresses += ("SMTP:$NewPrimarySMTPAddress")
      Set-Mailbox -Identity $user.Name -EmailAddresses $user.EmailAddresses

      #Modifica utente in exchange
      Set-MsolUserPrincipalName -UserPrincipalName $OldPrimarySMTPAddress -NewUserPrincipalName $NewPrimarySMTPAddress
    }

    ""; ""; Write-Host "Script terminato, verifica da console che tutto sia andato liscio! :-)" -f "green"; ""
  } catch {
    ""
    Write-Host "Errore nell'operazione, riprovare." -f "red"
    write-host $error[0]
    return ""
  }

}

# Start script
. Main
