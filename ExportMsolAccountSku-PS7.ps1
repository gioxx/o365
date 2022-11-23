<#
OFFICE 365: "User License Report" for PowerShell 7
---------------------------------------------------------------------------------------------------
Autore originale:	Kombaiah Murugan (8 feb. 2016)
URL originale:		http://www.morgantechspace.com/2016/02/get-all-licensed-office-365-users-with-powershell.html
Modifiche:        GSolone
Versione:         0.11
Utilizzo:         .\ExportMsolAccountSku-PS7.ps1
                  (opzionale, posizione CSV) .\ExportMsolAccountSku-PS7.ps1 -CSV C:\Licenze.csv
                  (opzionale, dominio da filtrare) .\ExportMsolAccountSku-PS7.ps1 -domain contoso.com
Info:				      https://gioxx.org/tag/o365-powershell
Fonti utilizzate:	https://theposhwolf.com/howtos/Set-MgUserLicense-PowerShell-Assign-O365-License/
                  https://o365reports.com/2021/11/23/office-365-license-reporting-and-management-using-powershell
Ultima modifica:	23-11-2022
Modifiche:
    0.11- correggo un errore causato dal caricamento modulo MSOnline prima di Graph (vedi https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/641).
          Approfitto di questo aggiornamento per correggere alcuni warning mostrati da VSCode e per modificare il blocco / funzione Pausa.
          Riordino le righe dei prodotti / licenze.
    0.10- revisione corrispondenza licenze e aggiungo POWERAPPS_DEV (Microsoft Power Apps for Developer).
    0.9- migro verso Microsoft Graph e PowerShell 7.
    0.8- aggiungo licenze per "Exchange Online (Piano 2)", "Enterprise Mobility + Security E3", "Dynamics 365 Team Members", "Power Automate Free", "Power Apps Plan 2 Trial", "Teams Exploratory" e "Microsoft Stream - Evaluation version" a quelle rilevate.
    0.7- aggiungo delimitatore ";" all'export-CSV.
    0.6- modifiche estetiche per parametri selezionati, modificata ricerca di dominio per includere anche i sottodomini. Includo la data odierna nel nome del file CSV estratto se non specificato diversamente da prompt.
    0.5 rev1- più che altro modifica estetica. Ho eliminato la colonna relativa all'avere licenza assegnata o meno, ho invertito il tipo di licenza posseduta prima dell'indirizzo di posta elettronica della persona.
    0.5- corretto errore nel nome del file CSV. Se non specificato, viene immediatamente forzato in una posizione standard (ho spostato l'istruzione relativa più in alto rispetto al box informativo mostrato ad avvio script). Aggiungo tra le licenze: MCOEV (Skype for Business Cloud PBX), PROJECTPROFESSIONAL (Project Online Professional) che sostituisce "PROJECTCLIENT" (il vecchio client Project Pro 2016 su PC) e DYN365_ENTERPRISE_PLAN1 (Dynamics 365 Plan 1 Enterprise Edition).
    Cambiato inoltre il metodo di estrazione dei dati: ora estraggo una riga per licenza, duplicando quindi l'assegnatario (ho comunque mantenuto anche il vecchio codice, nel caso in cui si preferisse una riga per utente, con tutte le licenze raggruppate).
    0.4- corretto errore nel "-f" quando si specifica un dominio di ricerca. Forzo un nome del CSV diverso se si specifica il dominio di ricerca.
    0.3- ho aggiunto il parametro UserPrincipalName all'estrazione, così da mostrare anche l'indirizzo di posta principale (generalmente corrispondente proprio a UserPrincipalName).
    0.2- includo la possibilità di filtrare un singolo dominio da riga di comando.
#>

# Verifica parametri da prompt
Param(
  [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)][string] $CSV,
  [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)][string] $Domain
)

# Se il CSV non è stato precedentemente specificato, utilizzo posizione e nome di default
if ([string]::IsNullOrEmpty($CSV)) {
  $DataOggi = Get-Date -format yyyyMMdd
  $CSV = "C:\temp\O365-User-License-Report_$DataOggi.csv"
}

#Cerco e carico i moduli necessari
$GraphModule = Get-Module -Name Microsoft.Graph -ListAvailable
if ($GraphModule.count -eq 0) {
  Write-Host "Installa il modulo Graph usando questo comando (poi rilancia questo script): `nInstall-Module Microsoft.Graph" -f "Yellow"
  Exit
}
else {
  Connect-MgGraph
}

$MOLModule=Get-Module -Name MSOnline -ListAvailable
if($MOLModule.count -eq 0) {
  Write-Host "Installa il modulo MSOnline usando questo comando (poi rilancia questo script): `nInstall-Module MSOnline" -f "Yellow"
  Exit
} else {
  Import-Module MSOnline -UseWindowsPowershell
  Connect-MsolService | Out-Null
}

# Main
""
Write-Host "        Office 365: User License Report for PowerShell 7 (Microsoft Graph)" -f "Green"
Write-Host "        ------------------------------------------"
Write-Host "         Lo script crea un report delle licenze assegnate agli utenti" -f "White"
Write-Host "         configurati sul server Exchange, salvando i risultati su un file CSV" -f "White"
if ([string]::IsNullOrEmpty($CSV) -eq $false) { Write-Host "[X]" -f "Yellow" -nonewline; }
Write-Host "         '" -f "White" -nonewline; Write-Host $CSV -f "Green" -nonewline; Write-Host "'" -f "White"
if ([string]::IsNullOrEmpty($CSV)) { Write-Host "[X]" -f "Yellow" -nonewline; }
Write-Host "         (rilancia lo script con parametro -CSV PERCORSOFILE.CSV per modificare)." -f "White"
Write-Host "         è possibile specificare un singolo dominio di ricerca ed esportazione da riga di comando." -f "White"
if ([string]::IsNullOrEmpty($Domain) -eq $false) { Write-Host "[X]" -f "Yellow" -nonewline; }
Write-Host "         (rilancia lo script con parametro -domain contoso.com per filtrare)." -f "White"
""

Write-Host -NoNewLine "Premi un tasto continuare (CTRL+C per annullare)";
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

""; Write-Progress -Activity "Download dati da Exchange" -Status "Scarico i dati relativi alle licenze, attendi."; ""

$ProcessedCount=0
if ([string]::IsNullOrEmpty($Domain)) {
  # DOMINIO NON SPECIFICATO: analizzo tutti gli utenti del server Exchange
  $users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" }
} else {
  # DOMINIO SPECIFICATO: lo esporto in maniera esclusiva se non specifico il CSV, il nome integra il dominio di ricerca
  if ([string]::IsNullOrEmpty($CSV)) {
    $CSV = "C:\temp\O365-User-License-Report_$($Domain).csv"
  }
  ""
  Write-Host "         Dominio specificato: " -nonewline; Write-Host "*$($Domain)" -f "Yellow"
  Write-Host "         CSV di destinazione: " -nonewline; Write-Host "$CSV" -f "Yellow"
  $users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" -and $_.UserPrincipalName -like "*@" + $Domain }
}

$userstotal = $users.Count
$users | Foreach-Object {
  $ProcessedCount++
  Write-Progress -Activity "Analisi in corso:" -Status "$ProcessedCount utenti di $userstotal" -PercentComplete (($ProcessedCount / $userstotal) * 100)
  $graphlicense = Get-MgUserLicenseDetail -UserId $_.UserPrincipalName
  ForEach ( $license in $($graphlicense.SkuPartNumber) ) {
    switch -wildcard ($($license)) {
      '*AAD_BASIC' { $licName = 'Azure Active Directory Basic' }
      '*AAD_PREMIUM_P1' { $licName = 'Azure Active Directory Premium P1' }
      '*AAD_PREMIUM_P2' { $licName = 'Azure Active Directory Premium P2' }
      '*AAD_PREMIUM' { $licName = 'Azure Active Directory Premium' }
      '*ADALLOM_O365' { $licName = 'Office 365 Advanced Security Management' }
      '*ADALLOM_S_O365' { $licName = 'POWER BI STANDALONE' }
      '*ADALLOM_S_STANDALONE' { $licName = 'Microsoft Cloud App Security' }
      '*ADALLOM_STANDALONE' { $licName = 'Microsoft Cloud App Security' }
      '*ATA' { $licName = 'Azure Advanced Threat Protection for Users' }
      '*ATP_ENTERPRISE_FACULTY' { $licName = 'Exchange Online Advanced Threat Protection' }
      '*ATP_ENTERPRISE' { $licName = 'Exchange Online Advanced Threat Protection' }
      '*BI_AZURE_P0' { $licName = 'Power BI (free)' }
      '*BI_AZURE_P1' { $licName = 'Power BI Reporting and Analytics' }
      '*BI_AZURE_P2' { $licName = 'Power BI Pro' }
      '*CCIBOTS_PRIVPREV_VIRAL' { $licName = 'Dynamics 365 AI for Customer Service Virtual Agents Viral SKU' }
      '*CRMINSTANCE' { $licName = 'Microsoft Dynamics CRM Online Additional Production Instance (Government Pricing)' }
      '*CRMIUR' { $licName = 'CRM for Partners' }
      '*CRMPLAN1' { $licName = 'Microsoft Dynamics CRM Online Essential (Government Pricing)' }
      '*CRMPLAN2' { $licName = 'Dynamics CRM Online Plan 2' }
      '*CRMSTANDARD' { $licName = 'CRM Online' }
      '*CRMSTORAGE' { $licName = 'Microsoft Dynamics CRM Online Additional Storage' }
      '*CRMTESTINSTANCE' { $licName = 'CRM Test Instance' }
      '*D365_CUSTOMER_SERVICE_ENT_ATTACH' { $licName = 'Dynamics 365 Sales Enterprise Edition' }
      '*D365_SALES_PRO' { $licName = 'Dynamics 365 for Sales Professional' }
      '*DESKLESS' { $licName = 'Microsoft StaffHub' }
      '*DESKLESSPACK_GOV' { $licName = 'Microsoft Office 365 (Plan K1) for Government' }
      '*DESKLESSPACK_YAMMER' { $licName = 'Office 365 Enterprise K1 with Yammer' }
      '*DESKLESSPACK' { $licName = 'Office 365 (Plan K1)' }
      '*DESKLESSWOFFPACK_GOV' { $licName = 'Microsoft Office 365 (Plan K2) for Government' }
      '*DESKLESSWOFFPACK' { $licName = 'Office 365 (Plan K2)' }
      '*DEVELOPERPACK_E5' { $licName = 'Microsoft 365 E5 Developer(without Windows and Audio Conferencing)' }
      '*DEVELOPERPACK' { $licName = 'Office 365 Enterprise E3 Developer' }
      '*DMENTERPRISE' { $licName = 'Microsoft Dynamics Marketing Online Enterprise' }
      '*DYN365_ENTERPRISE_CUSTOMER_SERVICE' { $licName = 'Dynamics 365 for Customer Service Enterprise Edition' }
      '*DYN365_ENTERPRISE_P1_IW' { $licName = 'Dynamics 365 P1 Trial for Information Workers' }
      '*DYN365_ENTERPRISE_PLAN1' { $licName = 'Dynamics 365 Plan 1 Enterprise Edition' }
      '*DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE' { $licName = 'Dynamics 365 for Sales and Customer Service Enterprise Edition' }
      '*DYN365_ENTERPRISE_SALES' { $licName = 'Dynamics 365 for Sales Enterprise Edition' }
      '*DYN365_ENTERPRISE_TEAM_MEMBERS' { $licName = 'Dynamics 365 for Team Members Enterprise Edition' }
      '*DYN365_FINANCIALS_BUSINESS_SKU' { $licName = 'Dynamics 365 for Financials Business Edition' }
      '*DYN365_MARKETING_APP' { $licName = 'Dynamics 365 Marketing' }
      '*DYN365_MARKETING_USER' { $licName = 'Dynamics 365 for Marketing USL' }
      '*DYN365_SALES_INSIGHTS' { $licName = 'Dynamics 365 AI for Sales' }
      '*DYN365_TEAM_MEMBERS' { $licName = 'Dynamics 365 Team Members' }
      '*Dynamics_365_for_Operations' { $licName = 'Dynamics 365 Unf Ops Plan Ent Edition' }
      '*ECAL_SERVICES' { $licName = 'ECAL' }
      '*EMS' { $licName = 'Enterprise Mobility + Security E3' }
      '*EMSPREMIUM' { $licName = 'Enterprise Mobility + Security E5' }
      '*ENTERPRISEPACK_B_PILOT' { $licName = 'Office 365 (Enterprise Preview)' }
      '*ENTERPRISEPACK_FACULTY' { $licName = 'Office 365 (Plan A3) for Faculty' }
      '*ENTERPRISEPACK_GOV' { $licName = 'Microsoft Office 365 (Plan G3) for Government' }
      '*ENTERPRISEPACK_STUDENT' { $licName = 'Office 365 (Plan A3) for Students' }
      '*ENTERPRISEPACK' { $licName = 'Office 365 Enterprise E3' }
      '*ENTERPRISEPACKLRG' { $licName = 'Office 365 Enterprise E3 LRG' }
      '*ENTERPRISEPACKWITHOUTPROPLUS' { $licName = 'Office 365 Enterprise E3 without ProPlus Add-on' }
      '*ENTERPRISEPREMIUM_NOPSTNCONF' { $licName = 'Enterprise E5 (without Audio Conferencing)' }
      '*ENTERPRISEPREMIUM' { $licName = 'Enterprise E5 (with Audio Conferencing)' }
      '*ENTERPRISEWITHSCAL_FACULTY' { $licName = 'Office 365 (Plan A4) for Faculty' }
      '*ENTERPRISEWITHSCAL_GOV' { $licName = 'Microsoft Office 365 (Plan G4) for Government' }
      '*ENTERPRISEWITHSCAL_STUDENT' { $licName = 'Office 365 (Plan A4) for Students' }
      '*ENTERPRISEWITHSCAL' { $licName = 'Office 365 Enterprise E4' }
      '*EOP_ENTERPRISE_FACULTY' { $licName = 'Exchange Online Protection for Faculty' }
      '*EOP_ENTERPRISE' { $licName = 'Exchange Online Protection' }
      '*EQUIVIO_ANALYTICS_FACULTY' { $licName = 'Office 365 Advanced Compliance for Faculty' }
      '*EQUIVIO_ANALYTICS' { $licName = 'Office 365 Advanced Compliance' }
      '*ESKLESSWOFFPACK_GOV' { $licName = 'Microsoft Office 365 (Plan K2) for Government' }
      '*EXCHANGE_ANALYTICS' { $licName = 'Microsoft MyAnalytics' }
      '*EXCHANGE_L_STANDARD' { $licName = 'Exchange Online (Plan 1)' }
      '*EXCHANGE_S_ARCHIVE_ADDON_GOV' { $licName = 'Exchange Online Archiving' }
      '*EXCHANGE_S_DESKLESS_GOV' { $licName = 'Exchange Kiosk' }
      '*EXCHANGE_S_DESKLESS' { $licName = 'Exchange Online Kiosk' }
      '*EXCHANGE_S_ENTERPRISE_GOV' { $licName = 'Exchange Plan 2G' }
      '*EXCHANGE_S_ENTERPRISE' { $licName = 'Exchange Online (Plan 2) Ent' }
      '*EXCHANGE_S_ESSENTIALS' { $licName = 'Exchange Online Essentials' }
      '*EXCHANGE_S_FOUNDATION' { $licName = 'Exchange Foundation for certain SKUs' }
      '*EXCHANGE_S_STANDARD_MIDMARKET' { $licName = 'Exchange Online Plan 1' }
      '*EXCHANGE_S_STANDARD' { $licName = 'Exchange Online Plan 2' }
      '*EXCHANGEARCHIVE_ADDON' { $licName = 'Exchange Online Archiving for Exchange Online' }
      '*EXCHANGEARCHIVE' { $licName = 'Exchange Online Archiving' }
      '*EXCHANGEDESKLESS' { $licName = 'Exchange Online Kiosk' }
      '*EXCHANGEENTERPRISE_FACULTY' { $licName = 'Exch Online Plan 2 for Faculty' }
      '*EXCHANGEENTERPRISE_GOV' { $licName = 'Microsoft Office 365 Exchange Online (Plan 2) only for Government' }
      '*EXCHANGEENTERPRISE' { $licName = 'Exchange Online Plan 2' }
      '*EXCHANGEESSENTIALS' { $licName = 'Exchange Online Essentials' }
      '*EXCHANGESTANDARD_GOV' { $licName = 'Microsoft Office 365 Exchange Online (Plan 1) only for Government' }
      '*EXCHANGESTANDARD_STUDENT' { $licName = 'Exchange Online (Plan 1) for Students' }
      '*EXCHANGESTANDARD' { $licName = 'Exchange Online Plan 1' }
      '*EXCHANGETELCO' { $licName = 'Exchange Online POP' }
      '*FLOW_FREE' { $licName = 'Microsoft Power Automate (Free)' }
      '*FLOW_O365_P2' { $licName = 'Flow for Office 365' }
      '*FLOW_O365_P3' { $licName = 'Flow for Office 365' }
      '*FLOW_P1' { $licName = 'Microsoft Flow Plan 1' }
      '*FLOW_P2' { $licName = 'Microsoft Flow Plan 2' }
      '*FORMS_PLAN_E3' { $licName = 'Microsoft Forms (Plan E3)' }
      '*FORMS_PLAN_E5' { $licName = 'Microsoft Forms (Plan E5)' }
      '*INFOPROTECTION_P2' { $licName = 'Azure Information Protection Premium P2' }
      '*INTUNE_A_VL' { $licName = 'Intune (Volume License)' }
      '*INTUNE_A' { $licName = 'Windows Intune Plan A' }
      '*INTUNE_O365' { $licName = 'Mobile Device Management for Office 365' }
      '*INTUNE_STORAGE' { $licName = 'Intune Extra Storage' }
      '*IT_ACADEMY_AD' { $licName = 'Microsoft Imagine Academy' }
      '*LITEPACK_P2' { $licName = 'Office 365 Small Business Premium' }
      '*LITEPACK' { $licName = 'Office 365 (Plan P1)' }
      '*LOCKBOX_ENTERPRISE' { $licName = 'Customer Lockbox' }
      '*LOCKBOX' { $licName = 'Customer Lockbox' }
      '*MCOCAP' { $licName = 'Command Area Phone' }
      '*MCOEV' { $licName = 'Skype for Business Cloud PBX' }
      '*MCOIMP' { $licName = 'Skype for Business Online (Plan 1)' }
      '*MCOLITE' { $licName = 'Lync Online (Plan 1)' }
      '*MCOMEETADV' { $licName = 'PSTN conferencing' }
      '*MCOPLUSCAL' { $licName = 'Skype for Business Plus CAL' }
      '*MCOPSTN1' { $licName = 'Skype for Business Pstn Domestic Calling' }
      '*MCOPSTN2' { $licName = 'Skype for Business Pstn Domestic and International Calling' }
      '*MCOSTANDARD_GOV' { $licName = 'Lync Plan 2G' }
      '*MCOSTANDARD_MIDMARKET' { $licName = 'Lync Online (Plan 1)' }
      '*MCOSTANDARD' { $licName = 'Skype for Business Online Standalone Plan 2' }
      '*MCVOICECONF' { $licName = 'Lync Online (Plan 3)' }
      '*MDM_SALES_COLLABORATION' { $licName = 'Microsoft Dynamics Marketing Sales Collaboration' }
      '*MEE_FACULTY' { $licName = 'Minecraft Education Edition Faculty' }
      '*MEE_STUDENT' { $licName = 'Minecraft Education Edition Student' }
      '*MEETING_ROOM' { $licName = 'Meeting Room' }
      '*MFA_PREMIUM' { $licName = 'Azure Multi-Factor Authentication' }
      '*MICROSOFT_BUSINESS_CENTER' { $licName = 'Microsoft Business Center' }
      '*MICROSOFT_REMOTE_ASSIST' { $licName = 'Dynamics 365 Remote Assist' }
      '*MIDSIZEPACK' { $licName = 'Office 365 Midsize Business' }
      '*MINECRAFT_EDUCATION_EDITION' { $licName = 'Minecraft Education Edition Faculty' }
      '*MS_TEAMS_IW' { $licName = 'Microsoft Teams' }
      '*MS-AZR-0145P' { $licName = 'Azure' }
      '*NBPOSTS' { $licName = 'Microsoft Social Engagement Additional 10k Posts (minimum 100 licenses) (Government Pricing)' }
      '*NBPROFESSIONALFORCRM' { $licName = 'Microsoft Social Listening Professional' }
      '*O365_BUSINESS_ESSENTIALS' { $licName = 'Microsoft 365 Business Basic' }
      '*O365_BUSINESS_PREMIUM' { $licName = 'Microsoft 365 Business Standard' }
      '*O365_BUSINESS' { $licName = 'Microsoft 365 Apps for business' }
      '*OFFICE_FORMS_PLAN_2' { $licName = 'Microsoft Forms (Plan 2)' }
      '*OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ' { $licName = 'Office ProPlus' }
      '*OFFICE365_MULTIGEO' { $licName = 'Multi-Geo Capabilities in Office 365' }
      '*OFFICESUBSCRIPTION_FACULTY' { $licName = 'Office 365 ProPlus for Faculty' }
      '*OFFICESUBSCRIPTION_GOV' { $licName = 'Office ProPlus' }
      '*OFFICESUBSCRIPTION_STUDENT' { $licName = 'Office ProPlus Student Benefit' }
      '*OFFICESUBSCRIPTION' { $licName = 'Microsoft 365 Apps for enterprise' }
      '*ONEDRIVESTANDARD' { $licName = 'OneDrive' }
      '*PAM_ENTERPRISE ' { $licName = 'Exchange Primary Active Manager' }
      '*PLANNERSTANDALONE' { $licName = 'Planner Standalone' }
      '*POWER_BI_ADDON' { $licName = 'Office 365 Power BI Addon' }
      '*POWER_BI_INDIVIDUAL_USE' { $licName = 'Power BI Individual User' }
      '*POWER_BI_INDIVIDUAL_USER' { $licName = 'Power BI for Office 365 Individual' }
      '*POWER_BI_PRO' { $licName = 'Power BI Pro' }
      '*POWER_BI_STANDALONE' { $licName = 'Power BI Standalone' }
      '*POWER_BI_STANDARD' { $licName = 'Power-BI Standard' }
      '*POWERAPPS_DEV' { $licName = 'Microsoft Power Apps for Developer' }
      '*POWERAPPS_INDIVIDUAL_USER' { $licName = 'Microsoft PowerApps and Logic flows' }
      '*POWERAPPS_O365_P2' { $licName = 'PowerApps' }
      '*POWERAPPS_O365_P3' { $licName = 'PowerApps for Office 365' }
      '*POWERAPPS_VIRAL' { $licName = 'Microsoft Power Apps (Plan 2 Trial)' }
      '*POWERFLOW_P1' { $licName = 'Microsoft PowerApps Plan 1' }
      '*POWERFLOW_P2' { $licName = 'Microsoft PowerApps Plan 2' }
      '*PREMIUM_ADMINDROID' { $licName = 'AdminDroid Office 365 Reporter' }
      '*PROJECT_CLIENT_SUBSCRIPTION' { $licName = 'Project Pro for Office 365' }
      '*PROJECT_ESSENTIALS' { $licName = 'Project Lite' }
      '*PROJECT_MADEIRA_PREVIEW_IW_SKU' { $licName = 'Dynamics 365 for Financials for IWs' }
      '*PROJECT_ONLINE_PRO' { $licName = 'Project Online Plan 3' }
      '*PROJECTCLIENT' { $licName = 'Project Professional' }
      '*PROJECTESSENTIALS' { $licName = 'Project Lite' }
      '*PROJECTONLINE_PLAN_1_FACULTY' { $licName = 'Project Online for Faculty Plan 1' }
      '*PROJECTONLINE_PLAN_1_STUDENT' { $licName = 'Project Online for Students Plan 1' }
      '*PROJECTONLINE_PLAN_1' { $licName = 'Project Online (Plan 1)' }
      '*PROJECTONLINE_PLAN_2_FACULTY' { $licName = 'Project Online for Faculty Plan 2' }
      '*PROJECTONLINE_PLAN_2_STUDENT' { $licName = 'Project Online for Students Plan 2' }
      '*PROJECTONLINE_PLAN_2' { $licName = 'Project Online and PRO' }
      '*PROJECTPREMIUM' { $licName = 'Project Online Premium' }
      '*PROJECTPROFESSIONAL' { $licName = 'Project Online Pro' }
      '*PROJECTWORKMANAGEMENT' { $licName = 'Office 365 Planner Preview' }
      '*RIGHTSMANAGEMENT_ADHOC' { $licName = 'Windows Azure Rights Management' }
      '*RIGHTSMANAGEMENT_STANDARD_FACULTY' { $licName = 'Azure Rights Management for faculty' }
      '*RIGHTSMANAGEMENT_STANDARD_STUDENT' { $licName = 'Information Rights Management for Students' }
      '*RIGHTSMANAGEMENT' { $licName = 'Azure Rights Management Premium' }
      '*RMS_S_ENTERPRISE_GOV' { $licName = 'Windows Azure Active Directory Rights Management' }
      '*RMS_S_ENTERPRISE' { $licName = 'Azure Active Directory Rights Management' }
      '*RMS_S_PREMIUM' { $licName = 'Azure Information Protection Plan 1' }
      '*RMS_S_PREMIUM2' { $licName = 'Azure Information Protection Premium P2' }
      '*SCHOOL_DATA_SYNC_P1' { $licName = 'School Data Sync (Plan 1)' }
      '*SHAREPOINT_PROJECT_EDU' { $licName = 'Project Online Service for Education' }
      '*SHAREPOINT_PROJECT' { $licName = 'SharePoint Online (Plan 2) Project' }
      '*SHAREPOINTDESKLESS_GOV' { $licName = 'SharePoint Online Kiosk' }
      '*SHAREPOINTDESKLESS' { $licName = 'SharePoint Online Kiosk' }
      '*SHAREPOINTENTERPRISE_EDU' { $licName = 'SharePoint Plan 2 for EDU' }
      '*SHAREPOINTENTERPRISE_GOV' { $licName = 'SharePoint Plan 2G' }
      '*SHAREPOINTENTERPRISE_MIDMARKET' { $licName = 'SharePoint Online (Plan 1)' }
      '*SHAREPOINTENTERPRISE' { $licName = 'SharePoint Online (Plan 2)' }
      '*SHAREPOINTLITE' { $licName = 'SharePoint Online (Plan 1)' }
      '*SHAREPOINTPARTNER' { $licName = 'SharePoint Online Partner Access' }
      '*SHAREPOINTSTANDARD_EDU' { $licName = 'SharePoint Plan 1 for EDU' }
      '*SHAREPOINTSTANDARD' { $licName = 'SharePoint Online Plan 1' }
      '*SHAREPOINTSTORAGE' { $licName = 'SharePoint Online Storage' }
      '*SHAREPOINTWAC_EDU' { $licName = 'Office Online for Education' }
      '*SHAREPOINTWAC_GOV' { $licName = 'Office Online for Government' }
      '*SHAREPOINTWAC' { $licName = 'Office Online' }
      '*SMB_APPS' { $licName = 'Business Apps (free)' }
      '*SMB_BUSINESS_ESSENTIALS' { $licName = 'Office 365 Business Essentials' }
      '*SMB_BUSINESS_PREMIUM' { $licName = 'Office 365 Business Premium' }
      '*SMB_BUSINESS' { $licName = 'Office 365 Business' }
      '*SPB' { $licName = 'Microsoft 365 Business' }
      '*SPE_E3' { $licName = 'Secure Productive Enterprise E3' }
      '*SPZA IW' { $licName = 'Microsoft PowerApps Plan 2 Trial' }
      '*SQL_IS_SSIM' { $licName = 'Power BI Information Services' }
      '*STANDARD_B_PILOT' { $licName = 'Office 365 (Small Business Preview)' }
      '*STANDARDPACK_FACULTY' { $licName = 'Office 365 (Plan A1) for Faculty' }
      '*STANDARDPACK_GOV' { $licName = 'Microsoft Office 365 (Plan G1) for Government' }
      '*STANDARDPACK_STUDENT' { $licName = 'Office 365 (Plan A1) for Students' }
      '*STANDARDPACK' { $licName = 'Office 365 (Plan E1)' }
      '*STANDARDWOFFPACK_FACULTY' { $licName = 'Office 365 Education E1 for Faculty' }
      '*STANDARDWOFFPACK_GOV' { $licName = 'Microsoft Office 365 (Plan G2) for Government' }
      '*STANDARDWOFFPACK_IW_FACULTY' { $licName = 'Office 365 Education for Faculty' }
      '*STANDARDWOFFPACK_IW_STUDENT' { $licName = 'Office 365 Education for Students' }
      '*STANDARDWOFFPACK_STUDENT' { $licName = 'Microsoft Office 365 (Plan A2) for Students' }
      '*STANDARDWOFFPACK' { $licName = 'Office 365 (Plan E2)' }
      '*STANDARDWOFFPACKPACK_FACULTY' { $licName = 'Office 365 (Plan A2) for Faculty' }
      '*STANDARDWOFFPACKPACK_STUDENT' { $licName = 'Office 365 (Plan A2) for Students' }
      '*STREAM_O365_E3' { $licName = 'Microsoft Stream for O365 E3 SKU' }
      '*STREAM_O365_E5' { $licName = 'Microsoft Stream for O365 E5 SKU' }
      '*STREAM' { $licName = 'Microsoft Stream' }
      '*SWAY' { $licName = 'Sway' }
      '*TEAMS_COMMERCIAL_TRIAL' { $licName = 'Microsoft Teams Commercial Cloud Trial' }
      '*TEAMS1' { $licName = 'Microsoft Teams' }
      '*THREAT_INTELLIGENCE' { $licName = 'Office 365 Threat Intelligence' }
      '*VIDEO_INTEROP ' { $licName = 'Skype Meeting Video Interop for Skype for Business' }
      '*VISIO_CLIENT_SUBSCRIPTION' { $licName = 'Visio Pro for Office 365' }
      '*VISIOCLIENT' { $licName = 'Visio Online Plan 2' }
      '*VISIOONLINE_PLAN1' { $licName = 'Visio Online Plan 1' }
      '*WACONEDRIVEENTERPRISE' { $licName = 'OneDrive for Business (Plan 2)' }
      '*WACONEDRIVESTANDARD' { $licName = 'OneDrive for Business with Office Online' }
      '*WACSHAREPOINTSTD' { $licName = 'Office Online STD' }
      '*WHITEBOARD_PLAN3' { $licName = 'White Board (Plan 3)' }
      '*WIN_DEF_ATP' { $licName = 'Windows Defender Advanced Threat Protection' }
      '*WIN10_PRO_ENT_SUB' { $licName = 'Windows 10 Enterprise E3' }
      '*WIN10_VDA_E3' { $licName = 'Windows E3' }
      '*WIN10_VDA_E5' { $licName = 'Windows E5' }
      '*WINDOWS_STORE' { $licName = 'Windows Store' }
      '*YAMMER_EDU' { $licName = 'Yammer for Academic' }
      '*YAMMER_ENTERPRISE_STANDALONE' { $licName = 'Yammer Enterprise' }
      '*YAMMER_ENTERPRISE' { $licName = 'Yammer for the Starship Enterprise' }
      '*YAMMER_MIDSIZE' { $licName = 'Yammer' }

      default { $licName = $license }
    }

    New-Object -TypeName PSObject -Property @{
      UserName=$_.DisplayName
      UserPrincipalName=$_.UserPrincipalName
      IsLicensed=$_.IsLicensed
      Licenses=$licName
    }
  }
} | Select-Object UserName,Licenses,UserPrincipalName | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"

""; Write-Host "Done." -f "Green"; ""

# Chiedo se visualizzare i risultati esportati nel file CSV
$message = "Devo aprire il file CSV $CSV ?"
$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Apri il file ora"
$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Non aprire il file adesso"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
$AproCSV = $host.ui.PromptForChoice("", $message, $options, 1)
if ($AproCSV -eq 0) { Invoke-Item $CSV }
