<#
OFFICE 365: Alias Management (Add or Remove mailbox aliases)
----------------------------------------------------------------------------------------------------------------
Autore:           GSolone
Versione:         0.1
Utilizzo:         .\AliasManagement.ps1
                  opzionale, utente: .\AliasManagement.ps1 mario.rossi mario.rossi.alias@contoso.com
                  opzionale, utente: .\AliasManagement.ps1 mario.rossi mario.rossi.alias@contoso.com -Remove
Info:				      https://gioxx.org/tag/o365-powershell
Ultima modifica:	30-04-2018
Modifiche:        -
#>

#Verifica parametri da prompt
Param(
  [Parameter(Position=0, Mandatory=$true, HelpMessage="Utente da modificare (es. mario.rossi)", ValueFromPipeline=$true)][string] $User,
  [Parameter(Position=1, Mandatory=$true, HelpMessage="Alias da aggiungere (es. mario.rossi.alias@contoso.com)", ValueFromPipeline=$true)][string] $Alias,
  [switch] $Remove
)

#Main
Function Main {
  ""; Write-Host "        Office 365: Alias Management" -f "Green"
  Write-Host "        ------------------------------------------"
  Write-Host "         Lo script aggiunge (o rimuove) un alias a una casella di posta elettronica"
  Write-Host "         è possibile lancare il PS1 passando da riga di comando l'utente da verificare"
  Write-Host "         es. .\AliasManagement.ps1 mario.rossi mario.rossi.alias@contoso.com"
  Write-Host "         (per rimuoverlo: .\AliasManagement.ps1 mario rossi mario.rossi.alias@contoso.com -Remove)"

  try {
    if ($Remove) {
      ""; Write-Host "Rimuovo $($Alias) da $User ed elenco gli indirizzi assegnati..." -f "Yellow";
      Set-Mailbox $User -EmailAddresses @{remove="$($Alias)"}
      Get-Recipient $User | Select Name -Expand EmailAddresses | where {$_ -like 'smtp*'}
      ""
    } else {
      ""; Write-Host "Aggiungo $($Alias) a $User ed elenco gli indirizzi assegnati..." -f "Yellow";
      Set-Mailbox $User -EmailAddresses @{add="$($Alias)"}
      Get-Recipient $User | Select Name -Expand EmailAddresses | where {$_ -like 'smtp*'}
      ""
    }
  }
  catch
  {
    Write-Host "Errore nell'operazione, riprovare." -f "Red"
    write-host $error[0]
    return ""
  }
}

# Start script
. Main
