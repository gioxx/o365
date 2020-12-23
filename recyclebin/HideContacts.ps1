<#
	OFFICE 365: "Hide a Contact or Mailbox"
	---------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.2
	Utilizzo:			.\HideContacts.ps1
						(opzionale, posizione Contact) .\HideContacts.ps1 -Contact mario.rossi
						(opzionale, posizione Mailbox) .\HideContacts.ps1 -Mailbox mario.rossi
	Info:				http://gioxx.org/tag/o365-powershell
	Ultima modifica:	08-09-2016
	Modifiche:
		0.2- includo la possibilità di passare il contatto da riga di comando. Ho aggiunto inoltre la possibilità di nascondere l'indirizzo di una mailbox, precedentemente non prevista.
#>

# Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Contact,
	[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Mailbox
)

""
Write-Host "        Office 365: Hide a Contact or Mailbox" -f "Green"
Write-Host "        ------------------------------------------"
Write-Host "         Lo script nasconde un indirizzo dalla rubrica di Exchange (GAL)." -f "White"
Write-Host "         Lanciarlo con il giusto parametro per nascondere il contatto:" -f "White"
Write-Host "          .\HideContacts.ps1 -Contact mario.rossi@contoso.com" -f "Yellow" -nonewline; Write-Host " (nasconde un contatto)" -f "White"
Write-Host "          .\HideContacts.ps1 -Mailbox mario.rossi@contoso.com" -f "Yellow" -nonewline; Write-Host " (nasconde l'indirizzo di una mailbox)" -f "White"
Write-Host "         (rilancia lo script il giusto parametro se non lo hai fatto)." -f "White"
""

if ([string]::IsNullOrEmpty($Contact) -eq $true) {
		# Contatto non specificato
		Write-Host "Contatto non specificato." -f "Red"
	} else {
		# Contatto specificato, procedo
		Write-Host "Contatto specificato: " -nonewline; Write-Host "$Contact" -f "Yellow" -nonewline; Write-Host ", procedo."
		Set-MailContact -Identity $Contact -HiddenFromAddressListsEnabled $true
		""
		Write-Host "Ho provato a nascondere $Contact dalla rubrica, verifico:" -f "Green"
		Get-MailContact -Identity $Contact | ft HiddenFromAddressListsEnabled
	}
	
if ([string]::IsNullOrEmpty($Mailbox) -eq $true) {
		# Mailbox non specificato
		Write-Host "Mailbox non specificata." -f "Red"
	} else {
		# Mailbox specificata, procedo
		Write-Host "Mailbox specificata: " -nonewline; Write-Host "$Mailbox" -f "Yellow" -nonewline; Write-Host ", procedo."
		Set-Mailbox -HiddenFromAddressListsEnabled $true -Identity $Mailbox
		""
		Write-Host "Ho provato a nascondere $Mailbox dalla rubrica, verifico:" -f "Green"
		Get-Mailbox -Identity $Mailbox | ft HiddenFromAddressListsEnabled
	}
	
""