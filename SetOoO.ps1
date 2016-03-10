############################################################################################################################
# OFFICE 365: Set User Out of Office
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.1
# Utilizzo:				.\SetOoO.ps1
#			(opzionale, passaggio dati da prompt) .\SetOoO.ps1 user@contoso.com -Status Disabled
#			(opzionale, passaggio dati da prompt) .\SetOoO.ps1 user@contoso.com -Int "Testo per interni" -Ext "Testo per esterni"
#			(opzionale, passaggio dati da prompt) .\SetOoO.ps1 user@contoso.com -Start "07/03/2016 08:00:00" -End "14/03/2016 08:00:00"
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		02-03-2016
# Modifiche:			-
############################################################################################################################

# Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $SourceMailbox, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Status, 
	[Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Int,
	[Parameter(Position=3, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Ext,
	[Parameter(Position=4, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Start,
	[Parameter(Position=5, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $End
)

# Creazione form di selezione calendario
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$objForm = New-Object Windows.Forms.Form
$objForm.Size = New-Object Drawing.Size @(200,190) 
$objForm.StartPosition = "CenterScreen"
$objForm.KeyPreview = $True

$objForm.Add_KeyDown({
	if ($_.KeyCode -eq "Enter") 
		{
			$dtmDate=$objCalendar.SelectionStart
			$objForm.Close()
		}
	})

$objForm.Add_KeyDown({
	if ($_.KeyCode -eq "Escape") 
		{
			$objForm.Close()
		}
	})

$objCalendar = New-Object System.Windows.Forms.MonthCalendar
$objCalendar.ShowTodayCircle = $True
$objCalendar.MaxSelectionCount = 1
$objForm.Controls.Add($objCalendar) 
$objForm.Topmost = $True

# Inizio script principale

""
Write-Host "        Office 365: Set User Out of Office" -f "Green"
Write-Host "        ------------------------------------------"

#######################################################################################################################
# VERIFICA E RICHIESTA DEI DETTAGLI MANCANTI DA PROMPT, LI RICHIEDO A VIDEO
#######################################################################################################################

# -Mailbox da modificare
if ( [string]::IsNullOrEmpty($SourceMailbox) )
{
	Write-Host "          ATTENZIONE:" -f "Red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -f "Red"
	Write-Host "          nei dati richiesti qui di seguito" -f "Red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	""
	$SourceMailbox = Read-Host "Utente da impostare in OoO (esempio: user@contoso.com)     "
}
# -Verifico se c'è richiesta di disabilitazione di OoO
if ( $Status -eq "Disabled" ) { 
		#Write-Host "Debug - Disabled: $SourceMailbox"
		Write-Host "          Disattivo Out of Office per $SourceMailbox ..." -f "Yellow"
		Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState Disabled
		Write-Host "          Done." -nonewline -f "Green"; Write-Host " Verificare eventuali errori."; ""
		Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox
		Break }
# -Messaggio per utenti presenti sullo stesso server Exchange
if ( [string]::IsNullOrEmpty($Int) ) { $Int = Read-Host "Messaggio di assenza per indirizzi interni (stesso server) " }
# -Messaggio per utenti esterni al server Exchange
if ( [string]::IsNullOrEmpty($Ext) ) {  $Ext = Read-Host "Messaggio di assenza per indirizzi esterni (diverso server)" }

# -Se non è specificato alcun messaggio, ne imposto uno io di default.
# -Se messaggio interno è specificato ma esterno no (o viceversa), l'altro diventa identico (principale: interno)
if ( [string]::IsNullOrEmpty($Int) -and [string]::IsNullOrEmpty($Ext) ) {
	$Int = "Sono assente per ferie e avrò limitato accesso alla mia mailbox.<br />
	Risponderò alla vostra comunicazione appena possibile.
	<br /><br />
	Buona giornata."
}
if ( [string]::IsNullOrEmpty($Int) -eq $false ) { 
	if ( [string]::IsNullOrEmpty($Ext) ) {
		$Ext = $Int
	}
}
if ( [string]::IsNullOrEmpty($Ext) -eq $false ) { 
	if ( [string]::IsNullOrEmpty($Int) ) {
		$Int = $Ext
	}
}


#######################################################################################################################
# ATTIVAZIONE / DISATTIVAZIONE OOO
#######################################################################################################################		

try
{
	""
	Write-Host "Verifica del periodo assenza di $SourceMailbox ..." -f "Yellow"
	""
	# Verifica periodo OoO
	$message = "Non è stato specificato un periodo di assenza, giusto così?"
	$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Continua senza specificare un periodo di assenza"
	$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Specifica ora un periodo di assenza"
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
	$PermissionType = $host.ui.PromptForChoice("", $message, $options, 0)
	if ($PermissionType -eq 1) {
		""
		Write-Host "Seleziona ora nel popup il primo giorno di ferie e premi invio" -f "Cyan"
		$objForm.Text = "Seleziona la data di inizio OoO (primo giorno di assenza)"
		$objForm.Add_Shown({$objForm.Activate()})  
		[void] $objForm.ShowDialog() 
		$Start=$dtmDate
		Write-Host "Seleziona ora nel popup l'ultimo giorno di ferie e premi invio" -f "Cyan"
		""
		$objForm.Text = "Seleziona la data di fine OoO (ultimo giorno di assenza)"
		$objForm.Add_Shown({$objForm.Activate()})  
		[void] $objForm.ShowDialog() 
		$End=$dtmDate
	}
	
	# Verifica flag di attivazione o disattivazione (default: Enabled).
	# Se il range date è specificato (solo se completo), il default passa a Scheduled.
	if ( [string]::IsNullOrEmpty($Status) ) { $Status = "Enabled" }
	if ( [string]::IsNullOrEmpty($Start) -eq $false -and [string]::IsNullOrEmpty($End) -eq $false ) { $Status = "Scheduled" }

	# Riepilogo dei dati pre-modifica
	""
	Write-Host "-- " -nonewline; Write-Host "Riepilogo dati" -nonewline -f "Green"; Write-Host " --"
	Write-Host "Casella da modificare      : " -nonewline -f "yellow"; Write-Host "$SourceMailbox"
	Write-Host "Status da impostare        : " -nonewline -f "yellow"; Write-Host "$Status"
	Write-Host "Messaggio di assenza (int) : " -nonewline -f "yellow"; Write-Host "$Int"
	Write-Host "Messaggio di assenza (est) : " -nonewline -f "yellow"; Write-Host "$Ext"
	# Mostro inizio-fine assenza solo se il periodo è interamente specificato,
	# altrimenti utilizzo un periodo esteso e manuale, senza scadenza.
	if ( [string]::IsNullOrEmpty($Start) -eq $false -and [string]::IsNullOrEmpty($End) -eq $false ) {
		Write-Host "Inizio assenza             : " -nonewline -f "yellow"; Write-Host "$Start"
		Write-Host "Termine assenza            : " -nonewline -f "yellow"; Write-Host "$End" 
	} else { 
		Write-Host "Periodo di assenza         : " -nonewline -f "yellow"; Write-Host "manuale, senza scadenza" 
		Write-Host "                             " -nonewline -f "yellow"; Write-Host "(Il periodo di assenza deve essere specificato interamente)" 
	}
	""
	
	# Imposto OoO
	Write-Host "Imposto Out of Office per $SourceMailbox ..." -f "Yellow"
	if ( $Status -eq "Enabled" ) { 
		#Write-Host "Debug - Enabled: $SourceMailbox"
		Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState Enabled -InternalMessage $Int -ExternalMessage $Ext }
	if ( $Status -eq "Scheduled" ) { 
		#Write-Host "Debug - Scheduled: $SourceMailbox"
		Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState Scheduled -StartTime $Start -EndTime $End -InternalMessage $Int -ExternalMessage $Ext }
	""; Write-Host "Done." -nonewline -f "Green"; Write-Host " Verificare eventuali errori.";
	Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox
}
catch
{
	Write-Host "Errore nell'operazione, riprovare." -f "Red"
	Write-Host $error[0]
	return ""
}