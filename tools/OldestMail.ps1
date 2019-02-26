<#	O365 PShell Snippet:	Get oldest message in a Mailbox
	Autore (ver.-mod.):		GSolone (0.1 ult.mod. 21/11/18)
	Utilizzo:				.\OldestMail.ps1 user@contoso.com
	Info:					https://gioxx.org/tag/o365-powershell
#>

# Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $SourceMailbox
)

Get-MailboxFolderStatistics -IncludeOldestAndNewestItems -Identity $SourceMailbox | 
    Where OldestItemReceivedDate -ne $null | 
    Sort OldestItemReceivedDate | 
    Select -First 1 OldestItemReceivedDate,FolderPath