<#	O365 PShell Snippet:	Set-MailboxRegionalConfiguration Locale IT
	Autore (ver.-mod.):		GSolone (0.3 ult.mod. 13/11/20)
	Utilizzo:				.\mboxRegionalConfigurationIT.ps1 user@contoso.com
	Info:					https://gioxx.org/tag/o365-powershell
#>

Param( 	[Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] [string] $SourceMailbox,
		[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] [string] $CSV
)
if ( [string]::IsNullOrEmpty($CSV) ) {
	if ( -not([string]::IsNullOrEmpty($SourceMailbox)) ) { 
		Write-Host "Modifico lingua della casella di posta $SourceMailbox " -f "yellow"
		Set-MailboxRegionalConfiguration $SourceMailbox -LocalizeDefaultFolderName:$true -Language it-IT
		Get-MailboxRegionalConfiguration $SourceMailbox
	}
} else {
	Import-CSV $CSV | foreach { 
		Write-Host "Modifico $_.EmailAddress" -f "Yellow"
		Set-MailboxRegionalConfiguration $_.EmailAddress -LocalizeDefaultFolderName:$true -Language it-IT
	}
}