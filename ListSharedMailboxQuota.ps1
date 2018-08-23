<#
	OFFICE 365: List Shared Mailbox Quota (GB)
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.1
	Utilizzo:			.\ListSharedMailboxQuota.ps1
	Info:				https://gioxx.org/tag/o365-powershell
	Fonti utilizzate:	https://blogs.technet.microsoft.com/heyscriptingguy/2013/02/27/get-exchange-online-mailbox-size-in-gb/
	Ultima modifica:	23-08-2018
	Modifiche:			-
#>

Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
  Get-MailboxStatistics |
  Select DisplayName, `
  @{name="TotalItemSize (GB)"; expression={[math]::Round( `
  ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}}, `
  ItemCount |
  Sort "TotalItemSize (GB)" -Descending