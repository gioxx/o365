#####################################################################
# O365 PShell Snippet:	Get Online Archive List						#
# Autore (ver.-mod.):	GSolone (0.1 ult.mod. 28/9/16)				#
# Utilizzo:				.\ListArchives.ps1							#
# Info:					http://gioxx.org/tag/o365-powershell		#
#####################################################################

""
Write-Host "Archivi presenti su Exchange:"
""
Write-Host "Utenti" -f "Yellow"
Get-Mailbox -ResultSize unlimited -Filter { ArchiveStatus -Eq "Active" -AND RecipientTypeDetails -eq 'UserMailbox'}
""; Write-Host "Shared Mailbox" -f "Yellow"; ""
Get-Mailbox -ResultSize unlimited -Filter { ArchiveStatus -Eq "Active" -AND RecipientTypeDetails -eq 'SharedMailbox'}