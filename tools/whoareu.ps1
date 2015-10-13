#####################################################################
# O365 PShell Snippet:	Who Are You?								#
# Autore (ver.-mod.):	GSolone (0.2 ult.mod. 27/8/15)				#
# Utilizzo:				.\whoareu.ps1 user@contoso.com				#
# Info:					http://gioxx.org/tag/o365-powershell		#
#####################################################################

$User=$args[0]
if ($args[0] -eq $null) { 
	""
	Write-Host "ATTENZIONE!" -f Yellow
	Write-Host "Per utilizzare lo script occorre richiamarlo passando l'utente. Es. whoareu mario.rossi"
	""
	} else { Get-User $User | ft Name,UserPrincipalName,WindowsLiveID,MicrosoftOnlineServicesID }