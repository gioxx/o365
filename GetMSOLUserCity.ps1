############################################################################################################################
# OFFICE 365: Get MSOL User City
#----------------------------------------------------------------------------------------------------------------
# Autore:				      GSolone
# Versione:				    0.1
# Utilizzo:				    .\GetMSOLUserCity.ps1
# Info:					      https://gioxx.org/tag/o365-powershell
# Ultima modifica:		18-11-2015
# Modifiche:
############################################################################################################################

#Verifica parametri da prompt
Param(
  [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)][string] $SearchCity
)

""; Write-Host "        Office 365: Get User City" -f "Green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -f "Red"
Write-Host "          L'operazione potrebbe durare diversi minuti (dipende dal numero di" -f "Red"
Write-Host "          utenti registrati all'interno del server Exchange)" -f "Red"; "";
Write-Host "-------------------------------------------------------------------------------------------------"; "";

if ( [string]::IsNullOrEmpty($SearchCity) )
{
  # Mancano i dettagli da prompt, li richiedo a video
  $SearchCity = Read-Host "Città da filtrare (esempio: Milano)"
  ""
}

try {
  Write-Progress -Activity "Download dati da Exchange" -Status "Cerco gli utenti che risultano essere in forza a $SearchCity , attendi..."
  Write-Host "$SearchCity :" -f "Yellow"
  Get-User -ResultSize Unlimited | where { $_.City -eq $SearchCity } | ft DisplayName, WindowsEmailAddress, Company, City
  Write-Host "All Done!" -f "Green"; "";
} catch {
  Write-Host "Errore nell'operazione, riprovare." -f "Red"
  Write-Host $error[0]
  return ""
}
