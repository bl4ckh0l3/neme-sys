<!-- #include virtual="/common/include/Objects/DBManagerClass.asp" -->
<!-- #include virtual="/common/include/Objects/LanguageClass.asp" -->
<!-- #include virtual="/common/include/Objects/ConfigClass.asp" -->
<!-- #include virtual="/common/include/Objects/objPageCache.asp"-->
<!-- #include virtual="/editor/include/InitData.inc" -->

<%
Dim publicDirVar, publicInstallDirVar, installFormPageVar, installResultPageVar, globalQueryFile, nemesiConfigFile

'************************************************************************************************************************************************************
'	LA VARIABILE SEGUENTE RAPPRESENTA L'UNICO PERCORSO FISICO CABLATO NELL'APPLICAZIONE;

'	SE LA DIRECTORY CON PERMESSO DI SCRITTURA MESSA A DISPOSIZIONE DAL VOSTRO PROVIDER FOSSE DIVERSA DA QUELLA PREFISSATA: "/public/*"
'	PRIMA DI AVVIARE LA PROCEDURA DI ISTALLAZIONE:
'	- MODIFICARE IL VALORE DELLA VARIABILE "publicDirVar" INDICANDO IL NUOVO PERCORSO
'	- SPOSTARE TUTTO IL CONTENUTO DELLA DIRECTORY /public/* DENTRO LA NUOVA DIRECTORY SCRIVIBILE;

'	UNA VOLTA GENERATO IL DATABASE TRAMITE LA PROCEDURA DI INSTALL, MODIFICARE DALLA CONSOLE DI AMMINISTRAZIONE IL CONF_VALUE DELLE KEYWORD:
'	dir_editor_upload;
'	dir_upload_news;
'	dir_upload_prod;
'	dir_upload_templ;
'	INDICANDO LA NUOVA DIRECTORY SCRIVIBILE PRINCIPALE, AL POSTO DI /public/*

publicDirVar = "/public"

'************************************************************************************************************************************************************

nemesiConfigFile = publicDirVar & "/conf/nemesi_config.xml"

	On Error Resume Next	
	Dim objConfig, strDbConnVar, installDirVar, strLineDbQuery, objFSO, queryFile, configFile, allAppVar
	

	'************* RECUPERO TUTTE LE QUERY NECESSARIE PER VALORIZZARE IL DATABASE E LE LANCIO IN SEQUENZA
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	Set objConfig = New ConfigClass	
	Set allAppVar = objConfig.getListaConfig()
	Set objConfig = nothing
	
	'************* CREO UN NUOVO FILE nemesi_config.xml dove inserisco tutte le variabili di configurazione dell'applicazione, da recuperare quando necessario
	Set configFile=objFSO.OpenTextFile(Server.MapPath(nemesiConfigFile), 2, True)
	configFile.writeLine("<config>")
		'configFile.writeLine("<currencysrvname srvname=""&Application("srt_default_server_name")&""><![CDATA["&Application("srt_default_server_name")&"]]></currencysrvname>")
		'configFile.writeLine("<strdbconn dbconn=""&Application("srt_dbconn")&""><![CDATA["&Application("srt_dbconn")&"]]></strdbconn>")
		
		for each x in allAppVar
			Set objConf = allAppVar(x)
			configFile.writeLine("<"&objConf.getKey()&" attr_"&objConf.getKey()&"="""&objConf.getValue()&"""></"&objConf.getKey()&">")
			Set objConf = nothing
		next
		
	configFile.writeLine("</config>")	
	configFile.Close
	Set configFile=Nothing
	Set allAppVar = nothing
	Set objFSO = nothing
	
	if(Err.number <> 0) then
		response.write(Err.description)
	else
		response.write("tutto OK")
	end if

%>

