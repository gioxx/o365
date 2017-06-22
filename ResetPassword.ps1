<#
	OFFICE 365: Reset User Password from PowerShell
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.2 rev1
	Utilizzo:			.\ResetPassword.ps1
						(opzionale, passaggio dati da prompt) .\ResetPassword.ps1 mario.rossi@contoso.com
	Info:				https://gioxx.org/tag/o365-powershell
	Ultima modifica:	22-06-2017
	Fonti utilizzate:	http://blogs.technet.com/b/heyscriptingguy/archive/2013/06/03/generating-a-new-password-with-windows-powershell.aspx
	Modifiche:
		0.2 rev1- banale modifica estetica, la scelta non ha un default (ho rimosso il "default: 1").
		0.2- consente il passaggio del dato di casella di posta da modificare (es. .\ResetPassword.ps1 mario.rossi@contoso.com
#>

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $RicercaUser
)

#Main
Function Main {

	#------------ Specifiche formattazione
	$Blocco = @{ForegroundColor="yellow"; object="          -----------------------------------------------------------";}
	$Blocco_Verde = @{ForegroundColor="green"; object="          -----------------------------------------------------------";}
	$Vuoto	= @{object="                                                                      ";}
	$VuotoN	= @{object="                                                          "; ; NoNewLine = $true;}
	$Apri 	= @{ForegroundColor="yellow"; object="         |"; NoNewLine = $true;}
	$Chiudi = @{ForegroundColor="yellow"; object=" |";}
	$Verde	= @{ForegroundColor="green"; NoNewLine = $true;}
	$Bianco	= @{ForegroundColor="white"; NoNewLine = $true;}
	#------------ Termine specifiche formattazione

	""
	Write-Host "        Office 365: User Password Reset" -f "green"
	Write-Host "        ------------------------------------------"
	
	if ( [string]::IsNullOrEmpty($RicercaUser) )
	{
		#MANCANO I DETTAGLI DA PROMPT, LI RICHIEDO A VIDEO
		Write-Host "          ATTENZIONE:" -foregroundcolor "red"
		Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
		Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
		""
		Write-Host "-------------------------------------------------------------------------------------------------"
		$RicercaUser = Read-Host "Utente (esempio: info@domain.tld) "
	} else {
		""
		Write-Host @Blocco; Write-Host @Vuoto
		Write-Host "          Modifica password utente di " @Bianco; Write-Host $RicercaUser @Verde
		Write-Host @Vuoto; Write-Host @Blocco
	}
	
	#------------ Blocco Operazione
	""
	Write-Host @Blocco
	Write-Host @Apri
	Write-Host " Scegli l'operazione da effettuare:                       " @bianco;
	Write-Host @Chiudi; Write-Host @Apri; Write-Host @VuotoN;
	Write-Host @Chiudi; Write-Host @Apri;
	Write-Host "    1- Cambio password con quella di default (Office2013) " @verde;
	Write-Host @Chiudi; Write-Host @Apri;
	Write-Host "       (Ricorda che in questo caso sarà obbligatorio      " @bianco;
	Write-Host @Chiudi; Write-Host @Apri;
	Write-Host "       cambiare la password al primo accesso!)            " @bianco;
	Write-Host @Chiudi; Write-Host @Apri; Write-Host @VuotoN;
	Write-Host @Chiudi; Write-Host @Apri;
	Write-Host "    2- Cambio password (generata randomicamente)          " @verde;
	Write-Host @Chiudi; Write-Host @Apri;
	Write-Host "    3- Cambio password (scelta dall'amministratore)       " @verde;
	Write-Host @Chiudi; Write-Host @Blocco
	""
	do {
		try {
			$numOk = $true
			[int]$ChangePasswd = Read-Host "Operazione scelta"
		} # end try
		catch {$numOK = $false}
		} # end do 
	until (($ChangePasswd -ge 1 -and $ChangePasswd -lt 4) -and $numOK)
	
	try
	{
		""
		#Cambio password
		switch ($ChangePasswd) 
			{ 
				1 	{
						#DEFAULT (Necessario cambio al primo accesso utente)
						Write-Host "Imposto password:" -f "Green"
						Set-MsolUserPassword -UserPrincipalName $RicercaUser -NewPassword "Office2013"
					}
				2 	{
						#GENERATA RANDOMICAMENTE
						$alphabet=$NULL;For ($a=65;$a –le 90;$a++) {$alphabet+=,[char][byte]$a }
						function GET-Temppassword() {
							Param(
							[int]$length=10,
							[string[]]$sourcedata
							)
							for ($loop=1; $loop –le $length; $loop++) {
								$TempPassword+=($sourcedata | GET-RANDOM)
							}
							return $TempPassword
						}
						Write-Host "Imposto password:" -f "Green"
						Set-MsolUserPassword -UserPrincipalName $RicercaUser -NewPassword $TempPassword -ForceChangePassword $false
					}
				3 	{
						#SCELTA DALL'AMMINISTRATORE
						$AdminRequest = Read-Host "Password (almeno 8 caratteri, almeno una maiuscola e un numero) "
						""
						Write-Host "Imposto password:" -f "green"
						Set-MsolUserPassword -UserPrincipalName $RicercaUser -NewPassword $AdminRequest -ForceChangePassword $false
					}
				
				#SCELTA DEFAULT NEL BLOCCO OPERAZIONE
				default { Set-MsolUserPassword -UserPrincipalName $RicercaUser -NewPassword "Office2013" }
			}
		""
		Write-Host "Salvo errori, il cambio password è andato a buon fine." @verde
		""; ""
		
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -f "red"
		write-host $error[0]
		return ""
	}
	
}

# Start script
. Main