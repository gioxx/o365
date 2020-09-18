<#
	OFFICE 365: Set Mailbox Forward
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.2
	Utilizzo:			.\SetMailboxForward-CSV.ps1 -CSV example.csv
	Info:				https://gioxx.org/tag/o365-powershell
	Fonti utilizzate:	https://docs.microsoft.com/it-it/exchange/recipients/user-mailboxes/email-forwarding?view=exchserver-2019
	Ultima modifica:	18-10-2019
	Modifiche:
		0.2- modifica minore, indirizzo_attuale diventa Source e indirizzo_forward diventa Target
#>

#Verifica parametri da prompt
Param(
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $CSV
)

#Main
Function Main {

	""; Write-Host "        Office 365: Set Mailbox Forward" -f "Green"
	Write-Host "        ------------------------------------------"
	Write-Host "         Costruire il file CSV con in colonna 1 l'attuale indirizzo di posta elettronica" -f "White"
	Write-Host "         e in colonna 2 il nuovo indirizzo verso il quale fare forward." -f "White"
	Write-Host "         Il titolo della prima colonna dovrà essere " -nonewline; Write-Host "Source" -f "Yellow" -nonewline; Write-Host ", la seconda " -f "White" -nonewline; Write-Host "Target" -f "Yellow"
	""
	Write-Host "         CSV di esempio:" -f "White"
	Write-Host "         Source,Target" -f "Gray"
	Write-Host "         test_1@contoso.onmicrosoft.com,test_1@contoso.com" -f "Gray"
	Write-Host "         test_2@contoso.onmicrosoft.com,test_2@contoso.com" -f "Gray"; ""
	
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
		Write-Progress -Activity "Download dati da Exchange" -Status "Ricerco le caselle elencate nel file CSV, attendi..."
		#Imposto il forward
		Import-Csv $CSV | ForEach-Object {
			Write-Progress -activity "Imposto forward" -status "$_.Source verso $_.Target"
			Set-Mailbox -Identity $_.Source -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $_.Target
			Get-Mailbox $_.Source | fl PrimarySmtpAddress,ForwardingSMTPAddress,DeliverToMailboxandForward
		}	
		Write-Host "Script terminato, verifica da console che tutto sia andato liscio! :-)" -f "Green"; ""
	}
	catch
	{
		""; Write-Host "Errore nell'operazione, riprovare." -f "Red"
		write-host $error[0]
		return ""
	}
}

# Start script
. Main