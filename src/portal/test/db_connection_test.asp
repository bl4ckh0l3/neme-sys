<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>test page</title>
</head>
<body>
<div id="warp">
<%
'************************************************************************************************************************************************************
dbuser="tosh.it"
dbpassword="Th.Sr19!"
dbname="tosh"
dbserver="webadmin.evolutionlab.it"


result_msg=""

On Error Resume Next	
	
'************* CREO LA NUOVA STRINGA DI CONNESSIONE CON I DATI FORNITI DALL'UTENTE
strDbConnVar ="driver={MySQL ODBC 3.51 Driver};uid="&dbuser&";pwd="&dbpassword&";database="&dbname&";Server="&dbserver&";port=3306"
'strDbConnVar ="driver={MySQL ODBC 3.51 Driver};uid=Sql198279;pwd=a34d7876;database=Sql198279_1;Server=62.149.150.77;port=3306"

'************* IMPOSTO L'OGGETTO objConn CON LA NUOVA STRINGA DI CONNESSIONE
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = strDbConnVar	
objConn.Open()
response.write("oggetto connection: "&typename(objConn)&" - stato connessione: "&objConn.state&"<br/><br/>")

'response.write("eseguo creazione tabella logs...<br/><br/>")

'************* CREO UNA TABELLA DI PROVA		
'strLineDbQuery = "CREATE TABLE IF NOT EXISTS `logs` (  `id` int(10) unsigned NOT NULL auto_increment,  `msg` TEXT default NULL,  `usr` varchar(50) NOT NULL,  `type` varchar(15) NOT NULL,  `date_event` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,  PRIMARY KEY  (`id`),  KEY `usr` (`usr`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8 ;"
'objConn.Execute(strLineDbQuery)

'response.write("eseguo query su tabella news")

'strLineDbQuery = "SELECT * FROM tosh_news ORDER BY tnws_ordine;"
'Set objRS = objConn.Execute(strLineDbQuery)
'if not(objRS.EOF) then						
'	do while not objRS.EOF
'		nome = objRS("tnwsd_nome")
'		testo = objRS("tnwsd_testo")
'		data = objRS("tnwsd_data") 
		
'		response.write(nome&"<br/>"&testo&"<br/><br/>"&data&"<br/><br/>")
'		objRS.moveNext()
'	loop				
'end if
'Set objRS = Nothing

objConn.close()
	
If Err.Number<>0 then
	result_msg=Err.description
end if

response.write("errori rilevati: "&result_msg&"<br/><br/>")
%>
</div>
</body>
</html>
