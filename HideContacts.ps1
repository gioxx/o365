############################################################################################################################
# OFFICE 365: Hide a Contact
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\HideContacts.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		07-05-2014
############################################################################################################################

""
Write-Host "        Office 365: Hide a contact" -foregroundcolor "green"
Write-Host "        ------------------------------------------"
Write-Host "          ATTENZIONE:" -foregroundcolor "red"
Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
""
Write-Host "-------------------------------------------------------------------------------------------------"

$result = 0
while ($result -eq 0) 
	{
		try
			{
				$ResetUser = Read-Host "Contatto (esempio: info@domain.tld)"
				Set-MailContact -Identity $ResetUser -HiddenFromAddressListsEnabled $true
				Write-Host "                                       $ResetUser nascosto dalla rubrica" -foregroundcolor "green"
				""
				""
				$title = ""
				$message = "Vuoi nascondere un ulteriore contatto?"
				$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Si"
				$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No"
				$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
				$result = $host.ui.PromptForChoice($title, $message, $options, 0)
			}
			catch
			{
				Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
				write-host $error[0]
				return ""
			}
	}