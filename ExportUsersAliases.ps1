<#
OFFICE 365: Export Recipient Aliases
----------------------------------------------------------------------------------------------------------------
Autore:					    GSolone
Versione:				    0.2
Utilizzo:				    .\ExportUsersAliases.ps1
                    obbligatorio, dominio di ricerca: .\ExportUsersAliases.ps1 contoso.com
                    opzionale, specifica posizione del file CSV esportato: .\ExportUsersAliases.ps1 -CSV C:\temp\Alias.csv
Info:					      https://gioxx.org/tag/o365-powershell
Fonti utilizzate:		https://social.technet.microsoft.com/Forums/exchange/en-US/a234ba3b-37b4-4333-8954-5f46885c5e20/how-to-list-email-addresses-and-aliases-for-each-user?forum=exchangesvrgenerallegacy
Ultima modifica:		06-06-2019
Modifiche:
    0.2- aggiungo delimitatore ";" all'export-CSV.
#>

<#
    Blocco originale:
    # Create an object to hold the results
    $addresses = @()
    # Get every mailbox in the Exchange Organisation
    $Mailboxes = Get-Mailbox -ResultSize Unlimited

    # Recurse through the mailboxes
    ForEach ($mbx in $Mailboxes) {
    # Recurse through every address assigned to the mailbox
    Foreach ($address in $mbx.EmailAddresses) {
    # If it starts with "SMTP:" then it's an email address. Record it
    if ($address.ToString().ToLower().StartsWith("smtp:")) {
    # This is an email address. Add it to the list
    $obj = "" | Select-Object Alias,EmailAddress
    $obj.Alias = $mbx.Alias
    $obj.EmailAddress = $address.ToString().SubString(5)
    $addresses += $obj
    }
    }
    }
    # Export the final object to a csv in the working directory
    $addresses | Export-Csv addresses.csv -NoTypeInformation -Delimiter ";"
#>

#Verifica parametri da prompt
Param(
  [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)][string] $RicercaDominio,
  [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)][string] $CSV
)

<#  Puoi modificare il valore $CSV per impostare un diverso nome del file CSV che verrà
    esportato dallo script (solo ciò che c'è tra le virgolette, ad esempio
    $CSV = "C:\temp\CSV.csv" (per modificare anche la cartella di esportazione),
    oppure $CSV = "Alias.csv" per salvare il file nella stessa cartella dello script. #>
$DataOggi = Get-Date -format yyyyMMdd
if ([string]::IsNullOrEmpty($CSV) -eq $true) {
  $CSV = "C:\temp\$($RicercaDominio)_$($DataOggi).csv"
}

""; Write-Host "        Office 365: Export Recipient Aliases" -f "Green"
Write-Host "        ------------------------------------------"
Write-Host "        CSV di destinazione: " -nonewline; Write-Host $CSV -f Yellow;

Write-Progress -Activity "Download dati da Exchange" -Status "Cerco e verifico gli alias assegnati a $RicercaDominio ..."
$addresses = @()
$Mailboxes = Get-Recipient -ResultSize Unlimited | where {$_.EmailAddresses -like "*@" + $RicercaDominio}
ForEach ($mbx in $Mailboxes) {
  Foreach ($address in $mbx.EmailAddresses) {
    if ($address.ToString().ToLower().EndsWith("@$RicercaDominio")) {
      $obj = "" | Select-Object DisplayName,PrimarySmtpAddress,RecipientType,EmailAddress
      $obj.DisplayName = $mbx.DisplayName
      $obj.PrimarySmtpAddress = $mbx.PrimarySmtpAddress
      $obj.RecipientType = $mbx.RecipientType
      $obj.EmailAddress = $address.ToString().SubString(5)
      $addresses += $obj
    }
  }
}
$addresses
""

# Chiedo se esportare i risultati in un file CSV
$message = "Devo esportare i risultati in $($CSV)?"
$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Esporta adesso i risultati"
$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non serve esportare i dati"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
$ExportCSV = $host.ui.PromptForChoice("", $message, $options, 1)
if ($ExportCSV -eq 0) {
  ""; Write-Host "Esporto l'elenco in $CSV e apro il file (salvo errori)." -f "Yellow"
  $addresses | Export-Csv $CSV -force -notypeinformation -Delimiter ";"
  Invoke-Item $CSV
}
