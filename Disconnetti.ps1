############################################################################################################################
# OFFICE 365: Disconnect from console
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\Disconnect.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		07-08-2014
# Modifiche:			-
############################################################################################################################

Remove-PSSession *
""
Write-Host "Connessione alla console terminata, ora è possibile uscire da Powershell" -foregroundcolor "green"
""