############################################################################################################################
# Powershell: Generate an encrypted password
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\genpasswd.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		07-08-2014
# Modifiche:			-
############################################################################################################################

clear
""
Write-Host "        Powershell: Generate an encrypted password" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -foregroundcolor "red"
Write-Host "          Ti verranno ora richiesti i dati di autenticazione alla console" -foregroundcolor "red"
Write-Host "          per poter criptare la tua password e generare il file txt." -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"

$credential = Get-Credential
if((Test-Path c:\temp) -eq 0) { new-item -type directory -path c:\temp }
$credential.Password | ConvertFrom-SecureString | Set-Content c:\temp\_PShellPasswd.txt
""
Write-Host "Operazione terminata, il tuo file è ora disponibile in c:\temp\_PShellPasswd.txt" -foregroundcolor "green"
Write-Host "Puoi spostarlo nella directory che preferisci e poi richiamarlo in fase di connessione alla Powershell" -foregroundcolor "green"
""