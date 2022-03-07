<#
OFFICE 365: Disconnect from PowerShell
----------------------------------------------------------------------------------------------------------------
Autore:           GSolone
Versione:			    0.1 rev1
Utilizzo:			    .\Disconnetti.ps1
Info:				      https://gioxx.org/tag/o365-powershell
Ultima modifica:  14-09-2017
Modifiche:
    0.1 rev1- modifiche estetiche minori.
#>

Remove-PSSession *
""; Write-Host "Connessione alla console terminata, ora è possibile chiudere la finestra di PowerShell" -f "Green"; "";
