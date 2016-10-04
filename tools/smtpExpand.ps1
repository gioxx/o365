#####################################################################
# O365 PShell Snippet:	Expand Mailbox SMTP Addresses          		#
# Autore (ver.-mod.):	GSolone (0.1 ult.mod. 15/7/16)				#
# Utilizzo:				.\smtpExpand.ps1 user@contoso.com			#
# Info:					http://gioxx.org/tag/o365-powershell		#
#####################################################################

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $SourceMailbox
)

""
Write-Host "Indirizzi di posta elettronica associati a $SourceMailbox " -f "yellow"
# Esclusioni applicate: NT AUTHORITY\SELF, S-1-5* (utenti non più presenti nel sistema)
Get-Recipient $SourceMailbox | Select Name -Expand EmailAddresses
""