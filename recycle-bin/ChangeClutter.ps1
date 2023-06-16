############################################################################################################################
# OFFICE 365: (Verify and) Change Clutter
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\ChangeClutter.ps1
#						opzionale, utente: .\ChangeClutter.ps1 mario.rossi@contoso.com
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		20-10-2015
# Modifiche:			-
############################################################################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $User
)

#Main
Function Main {
	
	""
	Write-Host "        Office 365: Change Clutter" -f "green"
	Write-Host "        ------------------------------------------"
	Write-Host "         Lo script verifica e permette di cambiare lo stato di Clutter di una casella di posta"
	Write-Host "         È possibile lancare il PS1 passando da riga di comando l'utente da verificare"
	Write-Host "         es. .\ChangeClutter.ps1 mario.rossi@contoso.com"
	
	if ([string]::IsNullOrEmpty($User) -eq $true)
	{
		#MANCANO I DETTAGLI DA PROMPT, LI RICHIEDO A VIDEO
		""
		Write-Host "-------------------------------------------------------------------------------------------------"
		""
		$User = Read-Host "Utente da verificare (esempio: mario.rossi@contoso.com)"
		""
	}
	
	try
	{
		
		""
		Write-Host "         Verifica stato Clutter di $User" -f "Yellow"
		Get-Clutter -Identity $User | ft isEnabled -AutoSize
		
		$title = ""
		$message = "Modificare lo stato di Clutter di $User ?"
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Disabilita", ""
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&Abilita", ""
		$Annulla = New-Object System.Management.Automation.Host.ChoiceDescription "A&nnulla", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No, $Annulla)
		$PermissionType = $host.ui.PromptForChoice($title, $message, $options, 0)
		if ($PermissionType -eq 0) { 
			""
			Write-Host "Hai scelto di disabilitare lo stato di Clutter di $User" -f "Yellow"
			Set-Clutter -Identity $User -Enable $false }
		if ($PermissionType -eq 1) { 
			""
			Write-Host "Hai scelto di abilitare lo stato di Clutter di $User" -f "Yellow"
			Set-Clutter -Identity $User -Enable $true }
		if ($PermissionType -eq 2) { 
			""
			Write-Host "Nessuna modifica da operare su $User" -f "Yellow" }
		
		Write-Host "Done." -f "Green"
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