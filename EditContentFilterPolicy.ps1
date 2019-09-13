<#
	OFFICE 365: Edit Hosted Content Filter Policy (WhitelistSender / WhitelistDomain / BlacklistSender / BlacklistDomain)
	----------------------------------------------------------------------------------------------------------------
	Autore:				GSolone
	Versione:			0.3
	Utilizzo:			.\EditContentFilterPolicy.ps1
						(opzionale, aggiungi utente Blacklist) .\EditContentFilterPolicy.ps1 Contoso -BlacklistSender user@contoso.com
						(opzionale, aggiungi dominio Blacklist) .\EditContentFilterPolicy.ps1 Contoso -BlacklistDomain contoso.com
						(opzionale, aggiungi utente Whitelist) .\EditContentFilterPolicy.ps1 Contoso -WhitelistSender user@contoso.com
						(opzionale, aggiungi dominio Whitelist) .\EditContentFilterPolicy.ps1 Contoso -WhitelistDomain contoso.com
						----------------------------------------------------------------------------------------------------------------
						(opzionale, rimuovi con un -Remove in chiusura) .\EditContentFilterPolicy.ps1 Contoso -WhitelistDomain contoso.com -Remove
	Info:				https://gioxx.org/tag/o365-powershell						
	Ultima modifica:	13-09-2019
	Modifiche:
		0.3- integro il nuovo ragionamento di Except che permette di escludere un utente dalla regola di blocco se fa parte di uno specifico gruppo.
		0.2 rev1- modifica estetica per correzione formattazione nel blocco informativo dello script.
		0.2- se non ci sono modifiche da operare, mostro le attuali impostazioni delle liste usando un Out-Gridview.
#>

#Verifica parametri da prompt
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] [string] $Spamfilter, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] [string] $BlacklistSender,
	[Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] [string] $BlacklistDomain,
	[Parameter(Position=3, Mandatory=$false, ValueFromPipeline=$true)] [string] $WhitelistSender,
	[Parameter(Position=4, Mandatory=$false, ValueFromPipeline=$true)] [string] $WhitelistDomain,
	[switch] $Remove
)

$ScriptName = $MyInvocation.MyCommand.Name
$AllowedSendersGroup = "allowedsenders@contoso.onmicrosoft.com"

""
Write-Host "        Office 365: Edit Hosted Content Filter Policy" -f "Green"
Write-Host "        ------------------------------------------";
Write-Host "        Aggiungere utenti in blacklist: .\$ScriptName Contoso -BlacklistSender user@contoso.com";
Write-Host "        Aggiungere domini in blacklist: .\$ScriptName Contoso -BlacklistDomain contoso.com";
Write-Host "        Aggiungere utenti in whitelist: .\$ScriptName Contoso -WhitelistSender user@contoso.com";
Write-Host "        Aggiungere domini in whitelist: .\$ScriptName Contoso -WhitelistDomain contoso.com"; "";
Write-Host "        Per rimuovere un utente o dominio da White/Blacklist, aggiungi -Remove in coda"; "        (es. .\$ScriptName Contoso -WhitelistDomain contoso.com -Remove)"; "";
$SpamFilters = Get-HostedContentFilterPolicy | Select Name
Write-Host "        Policy presenti:" $SpamFilters -f "Yellow"

if ( [string]::IsNullOrEmpty($Spamfilter) -eq $false ) {
	""; Write-Host "        Spam filter da analizzare: $($Spamfilter)"
	$StartEngine = 0
} else {
	""; $Spamfilter = Read-Host "        Spam filter da analizzare (esempio: @{Name=Contoso} --> Contoso)"
	$StartEngine = 0
}

try
{
	# Blacklist Sender
	if ( [string]::IsNullOrEmpty($BlacklistSender) -eq $false ) {
		""; Write-Host "        Blacklist Sender da modificare: $($BlacklistSender)"; "";
		if ($Remove) { Set-HostedContentFilterPolicy $Spamfilter -BlockedSenders @{remove="$BlacklistSender"} } else { Set-HostedContentFilterPolicy $Spamfilter -BlockedSenders @{add="$BlacklistSender"} }
		Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty BlockedSenders | ft Sender
		$StartEngine = 1
	}
	
	# Blacklist Domain
	if ( [string]::IsNullOrEmpty($BlacklistDomain) -eq $false ) {
		""; Write-Host "        Blacklist Domain da modificare: $($BlacklistDomain)"
		if ($Remove) { Set-HostedContentFilterPolicy $Spamfilter -BlockedSenderDomains @{remove="$BlacklistDomain"} } else { Set-HostedContentFilterPolicy $Spamfilter -BlockedSenderDomains @{add="$BlacklistDomain"} }
		Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty BlockedSenderDomains | ft Domain
		$StartEngine = 1
	}
	
	# Whitelist Sender
	if ( [string]::IsNullOrEmpty($WhitelistSender) -eq $false ) {
		""; Write-Host "        Whitelist Sender da modificare: $($WhitelistSender)"
		if ($Remove) { 
			Remove-DistributionGroupMember $AllowedSendersGroup -Member $WhitelistSender -confirm:$false
			Set-HostedContentFilterPolicy $Spamfilter -AllowedSenders @{remove="$WhitelistSender"}
		} else { 
			New-MailContact -DisplayName $WhitelistSender -Name $WhitelistSender -ExternalEmailAddress $WhitelistSender
			Set-MailContact $WhitelistSender -HiddenFromAddressListsEnabled $true
			$NewGroupMember = Add-DistributionGroupMember $AllowedSendersGroup -Member $WhitelistSender
			Set-HostedContentFilterPolicy $Spamfilter -AllowedSenders @{add="$WhitelistSender"}
		}
		Get-HostedContentFilterPolicy $Spamfilter | select -ExpandProperty AllowedSenders | ft Sender
		$StartEngine = 1
	}
	
	# Whitelist Domain
	if ( [string]::IsNullOrEmpty($WhitelistDomain) -eq $false ) {
		""; Write-Host "        Whitelist Domain da modificare: $($WhitelistDomain)"
		if ($Remove) { Set-HostedContentFilterPolicy $Spamfilter -AllowedSenderDomains @{remove="$WhitelistDomain"} } else { Set-HostedContentFilterPolicy $Spamfilter -AllowedSenderDomains @{add="$WhitelistDomain"} }
		Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty AllowedSenderDomains | ft Domain
		$StartEngine = 1
	}
	
	# Nessuna operazione eseguita
	if ($StartEngine -eq 0) { 
		""; Write-Host "        Nessuna modifica richiesta, ti mostro le black/whitelist della Spam Policy richiesta"; "";
		Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty BlockedSenders | Out-Gridview
		Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty BlockedSenderDomains | Out-Gridview
		Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty AllowedSenders | Out-Gridview
		Get-HostedContentFilterPolicy $Spamfilter | Select -ExpandProperty AllowedSenderDomains | Out-Gridview
		"";
	}
	
}
catch
{
	Write-Host "Errore nell'operazione, riprovare." -f "Red"
	Write-Host $error[0]
	return ""
}