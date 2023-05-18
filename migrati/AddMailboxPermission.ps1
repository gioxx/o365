<#
OFFICE 365: Add Mailbox Permission (Full Access / SendAs / GrantSendOnBehalfTo / Auto-mapping)
----------------------------------------------------------------------------------------------------------------
Autore:				    GSolone
Versione:			    0.13
Utilizzo:			    .\AddMailboxPermission.ps1
                  (opzionale, passaggio dati da prompt) .\AddMailboxPermission.ps1 shared@contoso.com mario.rossi@contoso.com (oppure .\AddMailboxPermission.ps1 shared@contoso.com mario.rossi@contoso.com)
Info:				      https://gioxx.org/tag/o365-powershell
Ultima modifica:	27-04-2018
Modifiche:
  0.13- ho aggiunto un Write-Host per notificare il GrantSendOnBehalfTo al giro di Get-MailboxPermission finale.
  0.12- ho aggiunto il GrantSendOnBehalfTo al giro di Get-MailboxPermission finale.
  0.11- ho solo modificato il Get-MailboxPermission finale.
  0.10- includo un blocco di verifica permessi finali (a operazione di ADD terminata) così da verificare gli utenti con accesso alla casella di posta (FullAccess e SendAs), escludendo NT AUTHORITY\SELF e S-1-5* (utenti non più presenti nel sistema).
  0.9- corretto variabile GiveAccessTo (riportata male nell'IF di controllo Empty String)
  0.8- corretto if-else di richiesta informazioni da prompt.
  0.7- lo script accetta adesso i parametri passati da riga di comando (-SourceMailbox e -GiveAccessTo)
  0.6- correzioni minori. Messo meglio in evidenza i dettagli riguardanti il "Send As" e il "Send on Behalf to".
  0.5- modificata la richiesta di GrantSendOnBehalfTo che ora viene mostrata solo se si rifiuta il SendAs.
  0.4- aggiungo la possibilità di specificare se l'utente deve inviare con proprietà Grantsendonbehalfto e non SendAs completo.
  0.3- prevedo la possibilità di scegliere l'auto-mapping della casella su Outlook, non utile nel caso di ShaRed Mailbox che impedirebbero in seguito la ricerca nelle sottocartelle.
  0.2- rev1/rev4-correpzioni minori, inclusa adesso la possibilità di modificare ulteriormente le ACL dando anche accesso "Invia Come" (SendAs).
#>

Param(
  [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)][string] $SourceMailbox,
  [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)][string] $GiveAccessTo
)

""; Write-Host "        Office 365: Add Mailbox Permission" -f "Green"
Write-Host "        ------------------------------------------"

if ( [string]::IsNullOrEmpty($SourceMailbox) -OR [string]::IsNullOrEmpty($GiveAccessTo) ) {
  # Mancano i dettagli da prompt, li richiedo a video
  Write-Host "          ATTENZIONE:" -f "Red"
  Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -f "Red"
  Write-Host "          nei dati richiesti qui di seguito" -f "Red"; ""
  Write-Host "-------------------------------------------------------------------------------------------------"; ""
  $SourceMailbox = Read-Host "Casella alla quale dare accesso (esempio: sharedmailbox@contoso.com)"
  $GiveAccessTo = Read-Host "Utente al quale dare Full Access (esempio: mario.rossi@contoso.com) "
  ""
}

try {
  ""; Write-Host "Autorizzo $GiveAccessTo a utilizzare $SourceMailbox ..." -f "Yellow"; ""
  $message = "$GiveAccessTo deve caricare automaticamente $SourceMailbox in Outlook (auto-mapping)?"
  $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
  $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
  $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
  $PermissionType = $host.ui.PromptForChoice("", $message, $options, 0)
  if ($PermissionType -eq 0) {
    Add-MailboxPermission -Identity $SourceMailbox -User $GiveAccessTo -AccessRights FullAccess
    ""; Write-Host "Accesso consentito a $GiveAccessTo (su $SourceMailbox), auto-mapping attivo" -f "Green"; ""
  } else {
    ""
    Add-MailboxPermission -Identity $SourceMailbox -User $GiveAccessTo -AccessRights FullAccess -AutoMapping:$false
    ""; Write-Host "Accesso consentito a $GiveAccessTo (su $SourceMailbox), auto-mapping DISATTIVATO"
    Write-Host "Ricordarsi di operare sull'Outlook dell'utente per caricare manualmente $SourceMailbox" -f "Green"; ""
  }

  $title = "SendAs - L'utente $GiveAccessTo deve poter inviare come fosse $SourceMailbox ?"
  $message = "ATTENZIONE: questo permette a $GiveAccessTo di scrivere a tutti gli effetti come fosse $SourceMailbox"
  $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
  $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
  $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
  $PermissionType = $host.ui.PromptForChoice($title, $message, $options, 0)
  if ($PermissionType -eq 0) {
    Add-RecipientPermission $SourceMailbox -Trustee $GiveAccessTo -AccessRights SendAs
    ""
  } else {
    ""; Write-Host "SendAs non impostato." -f "Red"
    $message = "Send On Behalf To - L'utente $GiveAccessTo deve poter almeno inviare a nome di $SourceMailbox (Es. Mario Rossi per conto di $SourceMailbox)?"
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
    $PermissionType = $host.ui.PromptForChoice("", $message, $options, 0)
    if ($PermissionType -eq 0) {
      Set-Mailbox $SourceMailbox -GrantSendOnBehalfTo @{add="$GiveAccessTo"}
      ""
    } else {
      ""; Write-Host "Send On Behalf To non impostato." -f "Red"
    }
  }

  ""; Write-Host "All Done!" -f "Green"
  Write-Host "Riepilogo accessi alla casella di $SourceMailbox " -f "yellow"
  # Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
  Get-MailboxPermission -Identity $SourceMailbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,AccessRights | ft User,AccessRights | out-string
  Get-RecipientPermission $SourceMailbox -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | ft Trustee, AccessRights | out-string
  Write-Host "GrantSendOnBehalfTo per $($SourceMailbox) (se presenti)" -f "Yellow"
  Get-Mailbox $SourceMailbox | Select -Expand GrantSendOnBehalfTo
  ""
} catch {
  Write-Host "Errore nell'operazione, riprovare." -f "Red"
  Write-Host $error[0]
  return ""
}
