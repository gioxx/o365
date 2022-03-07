<#	O365 PShell Snippet:	Shared Mailboxes: save a copy of all sent emails
Autore (ver.-mod.):				GSolone (0.1 ult.mod. 16/10/20)
Utilizzo:									.\SharedMailboxes-SaveSentEmails.ps1
Info:											https://gioxx.org/tag/o365-powershell
#>
$mboxCounter = 0
$Users = Get-Recipient -ResultSize Unlimited | Where {$_.RecipientTypeDetails -eq "SharedMailbox"}
Write-Host "Anteprima risultati" -f "Yellow"
$Users | ft DisplayName,Name,PrimarySmtpAddress -AutoSize
foreach ($mailbox in $Users) {
	$mboxCounter++
	$Percentage = [math]::Round($mboxCounter/$Users.Count*100,2)
	Write-Progress -Activity "Modifico comportamento Shared Mailbox ..." -Status "Casella: $($mailbox.PrimarySmtpAddress) ($($mboxCounter)/$($Users.Count) - $Percentage%)" -PercentComplete ($mboxCounter/$Users.Count*100)
	Set-Mailbox $mailbox.PrimarySmtpAddress -MessageCopyForSentAsEnabled $True
	Set-Mailbox $mailbox.PrimarySmtpAddress -MessageCopyForSendOnBehalfEnabled $True
}
