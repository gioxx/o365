<#
	OFFICE 365: List Mailboxes User Access
	-------------------------------------------------------------------------------------------------------------
	Autore:					GSolone
	Versione:				0.10
	Utilizzo:				.\ListMailboxesUserAccess-CSV.ps1
							(opzionale, posizione CSV) .\ListMailboxesUserAccess-CSV.ps1 -CSV C:\Utenti.csv
							(opzionale, filtro dominio) .\ListMailboxesUserAccess-CSV.ps1 -Domain contoso.com
							(opzionale, numero caselle da analizzare) .\ListMailboxesUserAccess-CSV.ps1 -Count 10
							(opzionale, analisi di tutte le caselle di posta) .\ListMailboxesUserAccess-CSV.ps1 -All
							(opzionale, analisi delle caselle in un CSV) .\ListMailboxesUserAccess-CSV.ps1 -Source C:\Mailbox.csv
	Info:					http://gioxx.org/tag/o365-powershell
	Ultima modifica:		29-04-2019
	Fonti utilizzate:		http://exchangeserverpro.com/list-users-access-exchange-mailboxes/
							http://mattellis.me/export-fullaccess-sendas-permissions-for-shared-mailboxes/
	Modifiche:
	0.10- la ricerca per dominio include i sottodomini.
	0.9- modifica da -Scope All a -All per analizzare tutte le caselle di posta.
	0.8- modifica estetica. Accanto alle opzioni attivate da riga di comando, propongo un "[X]" per darne immediato riscontro.
	0.7- ho aggiunto la possibilità di importare un file CSV con le caselle di posta da analizzare, per generare il report degli accessi FullAccess + SendAs. Il CSV dovrà contenere una colonna con gli indirizzi di posta, il titolo in testa dovrà essere "WindowsEmailAddress"
	0.6- ho completamente cambiato il metodo di ricerca e analisi dei permessi, basandomi sull'originale proposto da http://mattellis.me/export-fullaccess-sendas-permissions-for-shared-mailboxes/ e modificato per poter funzionare con Office 365 e PowerShell 2. Se il nome del file di output viene tenuto di default, aggiungo la data del giorno di estrazione.
	0.5- il default di analisi passa alle Shared Mailbox. Per allargare lo scope sarà necessario richiamare lo script con parametro -Scope All.
	0.4 rev2- aggiunta funzione di Pause per evitare di intercettare il tasto CTRL.
	0.4 (e rev1)- correzioni minori.
	0.3- introduco parametri da riga di comando per modificare posizione file CSV, filtrare un solo dominio di posta elettronica o limitare i risultati da ricercare.
	0.2- corretti problemi di ricerca ACL sulle caselle degli utenti di posta elettronica. Ora lo script ricerca tutti i permessi Full-Access impostati sulle caselle di posta elettronica di Exchange (a prescindere che si tratti di Shared Mailbox o caselle di posta personali).
	
	-------------------------------------------------------------------------------------------------------------
#>

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $CSV,
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Domain,
    [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Count,
	[Parameter(Position=3, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Source,
	[switch] $All
)

#Main
Function Main {

	################################################################################################
	# Puoi modificare il valore $ExportList per impostare un diverso nome del file CSV che verrà
	# esportato dallo script (solo ciò che c'è tra le virgolette, ad esempio
	# $ExportList = "C:\temp\export.csv" (per modificare anche la cartella di esportazione), 
	# oppure $ExportList = "Permessi.csv" per salvare il file nella stessa cartella dello script.
	# ATTENZIONE: utilizza (per comodità) nomi diversi nel caso in cui lo script esporti i permessi
	# 			  delle caselle Shared piuttosto che quelle personali.
	
	$DataOggi = Get-Date -format yyyyMMdd
	if ([string]::IsNullOrEmpty($CSV) -eq $true) {
		if ($All) {
			# Permessi di tutte le caselle
			$ExportList = "C:\temp\MailboxPermissions_$DataOggi.csv"
		} else { 
			# Permessi delle Shared
			$ExportList = "C:\temp\SharedMailboxPermissions_$DataOggi.csv" }
	} else { $ExportList = $CSV }
		
	################################################################################################

	""
	Write-Host "        Office 365: List Mailboxes User Access" -foregroundcolor "Green"
	Write-Host "        ------------------------------------------"
	Write-Host "         Lo script elenca i diritti Full Access per ogni casella di posta" -f "White"
	Write-Host "         presente su server Exchange (default: Shared Mailbox), salvando i risultati su un file CSV" -f "White"
	if ([string]::IsNullOrEmpty($CSV)) { Write-Host "[X]" -f "Yellow" -nonewline; }
	Write-Host "         '" -f "White" -nonewline; Write-Host $ExportList -f "Green" -nonewline; Write-Host "'" -f "White"
	if ([string]::IsNullOrEmpty($CSV) -eq $false) { Write-Host "[X]" -f "Yellow" -nonewline; }
	Write-Host "         (rilancia lo script con parametro -CSV PERCORSOFILE.CSV per modificare)." -f "White"
	if ([string]::IsNullOrEmpty($Source) -eq $false) { Write-Host "[X]" -f "Yellow" -nonewline; }
	Write-Host "         (rilancia lo script con parametro -Source per analizzare una lista di caselle salvata su file CSV)." -f "White"
	if ([string]::IsNullOrEmpty($Domain) -eq $false) { Write-Host "[X]" -f "Yellow" -nonewline; }
	Write-Host "         (rilancia lo script con parametro -Domain contoso.com per analizzare un singolo dominio)." -f "White"
	if ($All) { Write-Host "[X]" -f "Yellow" -nonewline; }
	Write-Host "         (rilancia lo script con parametro -All per analizzare tutte le caselle, non solo le Shared)." -f "White"
	""
	Write-Host "Trovi, vicino alle opzioni passate da riga di comando, il simbolo" -f "White" -nonewline; Write-Host " [X] " -f "Yellow" -nonewline; Write-Host "per indicare la relativa attivazione." -f "White"
	""
	
	<#do { $RicercaACL = Read-Host "Utente da cercare nei permessi (esempio: Mario Rossi) " } 
		while ($RicercaACL -eq [string]::empty)
	#>
	
	Function Pause($M="Premi un tasto continuare (CTRL+C per annullare)") {
		If($psISE) {
			$S=New-Object -ComObject "WScript.Shell"
			$B=$S.Popup("Fai clic su OK per continuare.",0,"In attesa dell'amministratore",0)
			Return
		}
		Write-Host -NoNewline $M;
		$I=16,17,18,20,91,92,93,144,145,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183;
		While($K.VirtualKeyCode -Eq $Null -Or $I -Contains $K.VirtualKeyCode) {
			$K=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}
	}
	Pause
	
	try
	{
		""; "";
		Write-Progress -Activity "Download dati da Exchange" -Status "Verifica permessi delle caselle di posta ed esportazione risultati in $ExportList, attendi..."
		
		<# VECCHIO SCRIPT - Ricerca di test da lanciare manualmente per verifiche
		Write-Progress -Activity "DEBUG QUERY" -Status "Se stai leggendo questo messaggio c'è qualche problema, modifica lo script per inserire nuovamente la query di Debug nella parte commentata"
		Get-Mailbox -ResultSize 100 | Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation TestQuery.csv
		#>
		
		if ([string]::IsNullOrEmpty($Count) -eq $true) { $Count = "Unlimited" }
		
		if ([string]::IsNullOrEmpty($Source) -eq $true) {
			# Campo Source non specificato, procedo con i parametri passati a riga di comando
			
			if ($All) {
				# Richiesta analisi di tutte le caselle di posta elettronica
				if ([string]::IsNullOrEmpty($Domain) -eq $false) {
					Write-Host "         Dominio da analizzare: *$($Domain)" -f "Yellow"
					Write-Host "         Counter analisi      : $Count" -f "Yellow"
					Write-Host "         Scope analisi        : All. Verranno analizzate tutte le caselle di posta." -f "Yellow"
					""
					# TUTTE LE CASELLE, DOMINIO -SPECIFICATO-
					# filtro solo il dominio richiesto (oltre le NT AUTHORITY\SELF e S-1-5*)
					"DisplayName" + "," + "PrimarySMTPAddress" + "," + "Full Access" + "," + "Send As" | Out-File $ExportList -Force
					$Mailboxes = Get-Mailbox -ResultSize $Count | where {$_.PrimarySmtpAddress -like "*" + $Domain} | Select Identity, PrimarySMTPAddress, DisplayName, DistinguishedName
					ForEach ($Mailbox in $Mailboxes) {
						$SendAs = Get-RecipientPermission $Mailbox.Identity -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | % {$_.Trustee}
						$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User}
						$Mailbox.DisplayName + "," + $Mailbox.PrimarySMTPAddress + "," + $FullAccess + "," + $SendAs | Out-File $ExportList -Append }
				} else {
					Write-Host "         Dominio da analizzare: All" -f "Yellow"
					Write-Host "         Counter analisi      : $Count" -f "Yellow"
					Write-Host "         Scope analisi        : All. Verranno analizzate tutte le caselle di posta." -f "Yellow"
					""
					# TUTTE LE CASELLE, DOMINIO -NON SPECIFICATO-
					# esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
					"DisplayName" + "," + "PrimarySMTPAddress" + "," + "Full Access" + "," + "Send As" | Out-File $ExportList -Force
					$Mailboxes = Get-Mailbox -ResultSize $Count | Select Identity, PrimarySMTPAddress, DisplayName, DistinguishedName
					ForEach ($Mailbox in $Mailboxes) {
						$SendAs = Get-RecipientPermission $Mailbox.Identity -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | % {$_.Trustee}
						$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User}
						$Mailbox.DisplayName + "," + $Mailbox.PrimarySMTPAddress + "," + $FullAccess + "," + $SendAs | Out-File $ExportList -Append }
					}
			} else {
				# Default: analisi Shared Mailbox
				if ([string]::IsNullOrEmpty($Domain) -eq $false) {
					Write-Host "         Dominio da analizzare: *$($Domain)" -f "Yellow"
					Write-Host "         Counter analisi      : $Count" -f "Yellow"
					Write-Host "         Scope analisi        : Default. Verranno analizzate le Shared Mailbox." -f "Yellow"
					""
					# SOLO -SHARED MAILBOX-, DOMINIO -SPECIFICATO-
					# filtro solo il dominio richiesto (oltre le NT AUTHORITY\SELF e S-1-5*)
					"DisplayName" + "," + "PrimarySMTPAddress" + "," + "Full Access" + "," + "Send As" | Out-File $ExportList -Force
					$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize $Count | where {$_.PrimarySmtpAddress -like "*" + $Domain} | Select Identity, PrimarySMTPAddress, DisplayName, DistinguishedName
					ForEach ($Mailbox in $Mailboxes) {
						$SendAs = Get-RecipientPermission $Mailbox.Identity -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | % {$_.Trustee}
						$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User}
						$Mailbox.DisplayName + "," + $Mailbox.PrimarySMTPAddress + "," + $FullAccess + "," + $SendAs | Out-File $ExportList -Append }
				} else {
					Write-Host "         Dominio da analizzare: All" -f "Yellow"
					Write-Host "         Counter analisi      : $Count" -f "Yellow"
					Write-Host "         Scope analisi        : Default. Verranno analizzate le Shared Mailbox." -f "Yellow"
					""
					# SOLO -SHARED MAILBOX-, DOMINIO -NON SPECIFICATO-
					# esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
					"DisplayName" + "," + "PrimarySMTPAddress" + "," + "Full Access" + "," + "Send As" | Out-File $ExportList -Force
					$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize $Count | Select Identity, PrimarySMTPAddress, DisplayName, DistinguishedName
					ForEach ($Mailbox in $Mailboxes) {
						$SendAs = Get-RecipientPermission $Mailbox.Identity -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | % {$_.Trustee}
						$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User}
						$Mailbox.DisplayName + "," + $Mailbox.PrimarySMTPAddress + "," + $FullAccess + "," + $SendAs | Out-File $ExportList -Append }
				}
			}
			
		} else {
			# Source non è vuoto, prendo in carico il file CSV e analizzo le caselle al suo interno
			if ([string]::IsNullOrEmpty($CSV) -eq $true) {
				$ExportList = "C:\temp\MailboxPermissions_$DataOggi.csv"
			} else { $ExportList = $CSV }
	
			Write-Host "         File CSV sorgente spefificato: $Source" -f "Yellow"
			Write-Host "         File CSV di destinazione     : $ExportList" -f "Yellow"
			""
			"DisplayName" + "," + "PrimarySMTPAddress" + "," + "Full Access" + "," + "Send As" | Out-File $ExportList -Force
			Import-Csv $Source | ForEach-Object {
				$Mailbox = Get-Mailbox $_.WindowsEmailAddress | Select Identity, PrimarySMTPAddress, DisplayName, DistinguishedName
				$SendAs = Get-RecipientPermission $Mailbox.Identity -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | % {$_.Trustee}
				$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User}
				$Mailbox.DisplayName + "," + $Mailbox.PrimarySMTPAddress + "," + $FullAccess + "," + $SendAs | Out-File $ExportList -Append
			}
			
		}
		
		Write-Host "Done." -f "Green"
		
		# Chiedo se visualizzare i risultati esportati nel file CSV
		$message = "Devo aprire il file CSV $ExportList ?"
		$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ""
		$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ""
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
		$Excel = $host.ui.PromptForChoice("", $message, $options, 1)
		if ($Excel -eq 0) { Invoke-Item $ExportList }
		""
		
	}
	catch
	{
		Write-Host "Errore nell'operazione, riprovare." -foregroundcolor "red"
		write-host $error[0]
		return ""
	}
}

# Start script
. Main