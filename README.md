Office 365 Powershell Tools
===================
Office 365 Powershell Tools è un set di script preparati per poter essere modificati / lanciati direttamente via PowerShell e amministrare più rapidamente un Exchange "in cloud"

> **In breve:**
> Gli script vengono distribuiti as-is, porre molta attenzione a ciò che si fa. Consiglio caldamente di leggere i dettagli degli script e delle loro funzioni, contenuti all'interno degli stessi. Se possibile, **effettuare dei test in ambiente NON di produzione**. Tutto questo è frutto di ore di lavoro, qualche imprecazione e molte ricerche. Molto difficilmente pubblicherò qualcosa che possa andare a spaccare il lavoro altrui su Exchange, ma è sempre bene verificare con un paio di occhi in più ciò che si va a modificare (e magari condividere l'esperienza, che male non fa mai!).

Nello specifico
--------
È possibile trovare tutti (o quasi) i riferimenti agli script nel mio [blog personale sotto apposito tag](http://gioxx.org/tag/o365-powershell). Gli articoli non sono stati scritti tutti, e alcuni script potrebbero non funzionare a dovere in diverso ambiente. Si parte dal presupposto che l'amministratore Exchange -*prima di lanciare qualsiasi script di questa cartella pubblica*- abbia già fatto connessione via Powershell al proprio server e abbia caricato i moduli MSOnline / MsolService:

	`$User = "esempio@domain.tld"`
	`$PWord = Get-Content C:\esempio\password.txt | ConvertTo-SecureString`
	`$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $PWord`
	`$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Credential -Authentication Basic -AllowRedirection`
	`Import-PSSession $Session`
	`$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $PWord`
	`Import-Module MSOnline`
	`Connect-MsolService -Credential $Credential`

Il codice sopra riportato non funzionerà nel caso in cui si faccia uso di autenticazione in due fattori. Per capire come collegarsi alla PowerShell con l'autenticazione 2-step attiva, consultare l'articolo pubblicato su [gioxx.org/2017/06/07/powershell-e-multi-factor-authentication-di-microsoft](https://gioxx.org/2017/06/07/powershell-e-multi-factor-authentication-di-microsoft/)

#### Dettagli degli script
Ciascuno script presente nella pagina contiene dei dettagli sul suo funzionamento, sulle modifiche e sulle eventuali fonti / documentazioni esterne consultate, nella porzione iniziale del codice del file ps1 (header).

Aprire il file e consultare le informazioni e le revisioni operate, in caso di difficoltà, è possibile contattarmi tramite [strumento interno di segnalazione problemi (Issues)](https://github.com/gioxx/o365/issues), e attendere una risposta / aggiornamento degli script.

----------

#### Note
Tutti gli script sono stati inizialmente sviluppati e verificati per connettersi e interagire con la versione 2.0 della PowerShell. È possibile forzare la connessione all'URL di Exchange Online puntando direttamente a https://ps.outlook.com/PowerShell-LiveID?PSVersion=2.0

Ogni script è stato poi provato e revisionato per funzionare anche con le nuove versioni di PowerShell, native su Windows 10:

    PS C:\>$PSVersionTable.PSVersion
    Major  Minor  Build  Revision
    -----  -----  -----  --------
    5      1      15063  502

Credits
-------
Qualche ringraziamento:

 - alle tante community sparse nel web, che si interessano all'argomento PowerShell e che permettono di imparare sempre cose nuove, quotidianamente,
 - a [GitHub](https://github.com/) per tutto ciò che mette a disposizione,
 - a [stackedit.io/editor](https://stackedit.io/editor) per l'ottimo editor MD online


----------
*ultima revisione: agosto 2017*
