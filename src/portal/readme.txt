Prerequisiti
=============
Assicurarsi che il webserver IIS sia istallato sul vostro sistema; versione 5 o superiore (è raccomandato l'uso della versione 6).
Assicurarsi che il Database MySql 5.0 o superiore sia istallato e configurato correttamente, e di avere a disposizione i seguenti dati:
- server database (ip o nome server);
- database username;
- database password;


Istallazione
=============
Per istallare correttametne neme-sys sul vostro sistema, eseguire i seguenti passaggi:

1) caricare sul vostro server, o sul vostro pc locale, nella directory designata per la vostra applicazione web, tutti i contenuti estratti dallo zip di istallazione;
2) aprire il browser all'url: http://www.vostro_sito.com/public/install/portalinstall.asp (se lavorate sul vostro pc locale, sostituire www.vostro_sito.com con localhost)
3) inserire i dati richiesti dalla pagina di istallazione;
4) al termine dell'istallazione sarete inviati direttamente alla pagina principale della console di amministrazione;
5) l'applicazione prevede di default che la directory http://www.vostro_sito.com/public/* sia scrivibile, per il caricamento dei template, allegati, immagini utente, immagini header flash, ecc;
    se per qualche ragione non fosse possibile utilizzare quella directory, spostare tutto il contenuto della directory http://www.vostro_sito.com/public/* nella nuova directory scrivibile e aggiornare il nuovo percorso 
    dalla console di amministrazione, nel menù "Configurazione portale",  modificando il valore delle seguenti variabili applicative:
	- dir_editor_upload;
	- dir_upload_news;
	- dir_upload_prod;
	- dir_upload_templ;
	- dir_upload_user;
	- dir_upload_header;


l'utente amministratore di default è:
user: administrator
pwd: admin


--
Il team neme-sys
www.neme-sys.it


********************************************************************************************************************************************************
********************************************************************************************************************************************************


Prerequisites
=============
You need to ensure that a IIS >= 5 or
higher (version 6 is recommended), is installed on your system.
You also need to ensure that MySQL >=5.0 or higher is installed and configured on your system.
Ensure you have this data:
- database server (ip address or server name);
- database username;
- database password;


Installations
=============
To install neme-sys  on your system, perform the following steps:

1) load on your server, or on your local PC, under the directory designated for your web application, all files and directories extracted from the installation .zip;
2) Open the browser at the URL: http://www.your_site.com/public/install/portalinstall.asp (if you work on your local computer, replace www.your_site.com with localhost)
3) Enter the required data on  the installation page;
4) After installation you will be sent directly to the main page of the administrative console;
5) The application set by default the directory http://www.your_site.com/public/* as writeable, it's the direcotry where the system load template, attachments, user imgs, header flash imgs, etc;
    if for any reason you cannot use this direcotry, move all contents of the directory http://www.your_site.com/public/* in to your new writable direcotry and modifiy from the admin console (menù Portal config), all this application variables value:
	- dir_editor_upload;
	- dir_upload_news;
	- dir_upload_prod;
	- dir_upload_templ;
	- dir_upload_user;
	- dir_upload_header;


default administrator user is:
user: administrator
pwd: admin

 
--
Your neme-sys-team
www.neme-sys.com