############################################################################################################################
# OFFICE 365: Bulk Set User Company Field
#----------------------------------------------------------------------------------------------------------------
# Autore:				GSolone
# Versione:				0.5
# Utilizzo:				.\SetCompanyBulk.ps1
# Info:					http://gioxx.org/tag/o365-powershell
# Ultima modifica:		13-10-2015
# Modifiche:	
#	0.5- Correzione minore: la ricerca viene effettuata sullo specifico dominio in ingresso, un eventuale sottodominio deve essere dichiarato (esempio: contoso.com nella ricerca non mostrerà i risultati di dep1.contoso.com nell'eventualità dep1 fosse un suo sottodominio).		
#	0.4- Modificato il $_.EmailAddresses in $_.PrimarySmtpAddress per mettere la Company in base all'indirizzo di posta principale e non considerare eventuali alias
#	0.3- Modificato il -ResultSize Unlimited per supportare il numero massimo di caselle
#	0.2- Inserita notifica di lavorazione per ciascun utente (prima assegnava il campo senza notificare alcunché durante la lavorazione, si arrivava direttamente alla fine del ciclo ForEach)
############################################################################################################################

#Main
Function Main {

	""
	Write-Host "        Office 365: Bulk Set User Company Field" -foregroundcolor "green"
	Write-Host "        ------------------------------------------"
	Write-Host "          ATTENZIONE:" -foregroundcolor "red"
	Write-Host "          Fare molta attenzione ai possibili errori di digitazione" -foregroundcolor "red"
	Write-Host "          nei dati richiesti qui di seguito" -foregroundcolor "red"
	""
	Write-Host "-------------------------------------------------------------------------------------------------"
	$RicercaDominio = Read-Host "Dominio da analizzare (esempio: domain.tld) "
	$RicercaCompany = Read-Host "Valore Company (esempio: Contoso S.r.l.)    "
	
	try
	{
		""
		Write-Host "Ricerco le caselle con il dominio che mi hai richiesto, attendi." -foregroundcolor "yellow"
		$RicercaMailbox= Get-Mailbox -ResultSize Unlimited | where {$_.PrimarySmtpAddress -like "*@" + $RicercaDominio}
		#$RicercaMailbox= Get-Mailbox -ResultSize Unlimited | where {$_.EmailAddresses -like "*" + $RicercaDominio}
		Write-Host "Applico il valore Company alle utenze rilevate ..." -foregroundcolor "yellow"
		""
		$RicercaMailbox | ForEach {Write-Host "Utente: $_"; Set-User $_.UserPrincipalName -Company $RicercaCompany}
		""
		Write-Host "Done." -foregroundcolor "green"
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