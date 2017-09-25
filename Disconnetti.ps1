<#
	OFFICE 365: Disconnect from console
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.1 rev1
	Utilizzo:			.\Disconnect.ps1
	Info:				http://gioxx.org/tag/o365-powershell
	Ultima modifica:	14-09-2017
	Modifiche:
		0.1 rev1- modifiche estetiche minori.
#>

Remove-PSSession *
""; Write-Host "Connessione alla console terminata, ora è possibile chiudere la finestra di PowerShell" -f "Green"; "";