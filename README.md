Office 365 Powershell Tools
===================
Office 365 Powershell Tools è un set di script preparati per poter essere modificati / lanciati direttamente via PowerShell e amministrare più rapidamente il vostro Exchange "in cloud"

Nello specifico
--------
Trovate tutti (o quasi) i riferimenti agli script nel mio [blog personale sotto apposito tag](http://gioxx.org/tag/o365-powershell). Gli articoli non sono tutti pronti e alcuni script potrebbero quindi non funzionare a dovere. Si parte dal presupposto che l'amministratore Exchange -prima di lanciare qualsiasi script di questa cartella pubblica- abbia già fatto connessione via Powershell al proprio server e abbia caricato i moduli MSOnline / MsolService:

	`$User = "esempio@domain.tld"`
	`$PWord = Get-Content C:\esempio\password.txt | ConvertTo-SecureString`
	`$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $PWord`
	`$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Credential -Authentication Basic -AllowRedirection`
	`Import-PSSession $Session`
	`$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $PWord`
	`Import-Module MSOnline`
	`Connect-MsolService -Credential $Credential`

> **Attenzione:**
> Gli script vengono distribuiti as-is, occhio a ciò che fate, soprattutto in ambiente di produzione. Vi consiglio caldamente di attendere i relativi articoli sul blog che veranno poi aggiunti a questo readme prima di utilizzare uno script. Se possibile **effettuate dei test in ambiente NON di produzione**. Tutto questo è frutto di lavoro, qualche imprecazione e molte ricerche, molto difficilmente pubblicherò qualcosa che possa andare a spaccare il vostro lavoro su Exchange ma è sempre bene verificare con un paio di occhi in più ciò che si va a toccare (e magari condividere l'esperienza, che male non fa mai!).


#### Dettagli degli script
Ciascuno script presente nella pagina contiene dei dettagli sul suo funzionamento, sulle modifiche e su fonti / documentazioni esterne utilizzate nella porzione iniziale del file ps1.

Aprite il file e consultate le informazioni e le revisioni operate, in caso di difficoltà vi invito a contattarmi tramite [strumento interno di segnalazione problemi (Issues)](https://github.com/gioxx/o365/issues) e attendere una risposta / aggiornamento degli script.

----------

#### Note
Tutti gli script sono stati sviluppati e verificati per connettersi e interagire con la versione 2.0 della PowerShell. Potete decidere di forzare la connessione al vostro Exchange Online puntando direttamente a https://ps.outlook.com/PowerShell-LiveID?PSVersion=2.0

ult.agg. aprile 2016