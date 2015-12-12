<%
Response.ContentType="text/html"
Response.Charset="UTF-8"
%>
<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/common/include/Objects/CacheClass.asp" -->

<%
Public Function getListaTamponi()
	on error resume next
	Dim objDB, strSQL, objRS, objConn, objDict, objCategoria
	getListaTamponi = null		
	strSQL = " SELECT"
	strSQL = strSQL & " tosh_tamponi.`id`,"
	strSQL = strSQL & " tosh_tamponi.`nome`,"
	strSQL = strSQL & " tosh_tamponi.`tipo_num`,"
	strSQL = strSQL & " tosh_tamponi_categorie.`ttci_desc` as description,"
	strSQL = strSQL & " tosh_tamponi_categorie.`ttci_lng` as lang,"
	strSQL = strSQL & " tosh_tamponi.`d_stamp`,"
	strSQL = strSQL & " tosh_tamponi.`b_stamp`,"
	strSQL = strSQL & " tosh_tamponi.`a_stamp`,"
	strSQL = strSQL & " tosh_tamponi.`d_tamp`,"
	strSQL = strSQL & " tosh_tamponi.`b_tamp`,"
	strSQL = strSQL & " tosh_tamponi.`p_tamp`,"
	strSQL = strSQL & " tosh_tamponi.`a_tamp`,"
	strSQL = strSQL & " tosh_tamponi.`di_gomma`,"
	strSQL = strSQL & " tosh_tamponi.`ds_gomma`,"
	strSQL = strSQL & " tosh_tamponi.`bi_gomma`,"
	strSQL = strSQL & " tosh_tamponi.`bs_gomma` ,"
	strSQL = strSQL & " tosh_tamponi.`pi_gomma`,"
	strSQL = strSQL & " tosh_tamponi.`ps_gomma`,"
	strSQL = strSQL & " tosh_tamponi.`path_foto`,"
	strSQL = strSQL & " tosh_tamponi.`path_schem`"
	strSQL = strSQL & " FROM `tosh_tamponi_categorie`"
	strSQL = strSQL & " RIGHT JOIN `tosh_tamponi` ON tosh_tamponi_categorie.ttci_tipo_id =tosh_tamponi.tipo_num;"
	
	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()		
	Set objRS = objConn.Execute(strSQL)
	if not(objRS.EOF) then			
		Set objDict = Server.CreateObject("Scripting.Dictionary")			
		do while not objRS.EOF
			Set objTampone = Server.CreateObject("Scripting.Dictionary")
			strID = objRS("id")
			tipo_numero = objRS("tipo_num")
			titolo_forma = tipo_numero&" "&objRS("description")		
			objTampone.add "id-tampone", strID
			objTampone.add "codice-tampone", ""&objRS("nome")
			objTampone.add "titolo-tampone", titolo_forma			
			objTampone.add "forma-tampone", "frontend.template_tamponi.field.label.forma_tampone_"&tipo_numero
			objTampone.add "lang-tampone", ""&objRS("lang")
			objTampone.add "area-stampa-tampone", Cint(objRS("b_stamp"))&" x "&Cint(objRS("a_stamp"))
			objTampone.add "diametro-stampa-tampone", ""&objRS("d_stamp")
			objTampone.add "base-tampone", objRS("b_tamp")&" x "&objRS("p_tamp")
			objTampone.add "altezza-tampone", ""&objRS("a_tamp")
			objTampone.add "diam-tampone", ""&objRS("d_tamp")
			objTampone.add "dim-gomma-A", ""&objRS("bi_gomma")	
			objTampone.add "dim-gomma-B", ""&objRS("bs_gomma")	
			objTampone.add "dim-gomma-C", ""&objRS("pi_gomma")	
			objTampone.add "dim-gomma-D", ""&objRS("ps_gomma")		
			objTampone.add "dim-gomma-E", ""&objRS("di_gomma")	
			objTampone.add "dim-gomma-F", ""&objRS("ds_gomma")	
			objTampone.add "path-small-img", ""&objRS("path_schem")	
			objTampone.add "path-medium-img", ""&objRS("path_foto")		

			'response.write("id-tampone:" 							&objTampone.item("id-tampone") 							&"<br>")
			'response.write("codice-tampone:" 					&objTampone.item("codice-tampone") 					&"<br>")
			'response.write("titolo-tampone:" 					&objTampone.item("titolo-tampone") 					&"<br>")			
			'response.write("forma-tampone:" 					&objTampone.item("forma-tampone") 					&"<br>")
			'response.write("lang-tampone:" 						&objTampone.item("lang-tampone") 						&"<br>")
			'response.write("area-stampa-tampone:" 		&objTampone.item("area-stampa-tampone") 		&"<br>")
			'response.write("diametro-stampa-tampone:" &objTampone.item("diametro-stampa-tampone") &"<br>")
			'response.write("base-tampone:" 						&objTampone.item("base-tampone") 						&"<br>")
			'response.write("altezza-tampone:" 				&objTampone.item("altezza-tampone") 				&"<br>")
			'response.write("diam-tampone:" 						&objTampone.item("diam-tampone") 						&"<br>")
			'response.write("dim-gomma-A:" 						&objTampone.item("dim-gomma-A") 						&"<br>")	
			'response.write("dim-gomma-B:" 						&objTampone.item("dim-gomma-B") 						&"<br>")	
			'response.write("dim-gomma-C:" 						&objTampone.item("dim-gomma-C") 						&"<br>")	
			'response.write("dim-gomma-D:" 						&objTampone.item("dim-gomma-D") 						&"<br>")		
			'response.write("dim-gomma-E:" 						&objTampone.item("dim-gomma-E") 						&"<br>")	
			'response.write("dim-gomma-F:" 						&objTampone.item("dim-gomma-F") 						&"<br>")	
			'response.write("path-small-img:" 					&objTampone.item("path-small-img") 					&"<br>")
			'response.write("path-medium-img:" 				&objTampone.item("path-medium-img") 				&"<br>")

			objDict.add objTampone, ""
			Set objTampone = nothing
			objRS.moveNext()
		loop
						
		Set getListaTamponi = objDict			
		Set objDict = nothing				
	end if
	
	Set objRS = Nothing
	Set objDB = Nothing
	
	if Err.number <> 0 then
		'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if		
End Function


if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if

	redirectPage = Application("baseroot")&"/editor/contenuti/ListaNews.asp"		
						
	Dim objNews
	Set objNews = New NewsClass
	
	Dim objLogger
	Set objLogger = New LogClass

	Dim objContentField
	Set objContentField = new ContentFieldClass			

	'/**
	'* news da inserire e recupero Max(ID) 
	'*/	
	Dim newMaxID

	
	Set objListTamponi = getListaTamponi()		

	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()	
	objConn.BeginTrans	
	
	for each x in objListTamponi
		newMaxID = objNews.insertNews(x("titolo-tampone"), null, null, null, null, "lista-tamponi", Now(), Now(), null, 1, null, null, null, objConn)		
		call objLogger.write("inserito tampone --> id: "&newMaxID&"; titolo: "&x.item("titolo-tampone"), "system", "debug")
		
		value = x.item("codice-tampone")
		call objContentField.insertFieldMatch(1, newMaxID, value, objConn)		
		value = x.item("forma-tampone")
		call objContentField.insertFieldMatch(2, newMaxID, value, objConn)
		value = x.item("area-stampa-tampone")
		call objContentField.insertFieldMatch(3, newMaxID, value, objConn)
		value = x.item("diametro-stampa-tampone")
		call objContentField.insertFieldMatch(4, newMaxID, value, objConn)
		value = x.item("altezza-tampone")
		call objContentField.insertFieldMatch(5, newMaxID, value, objConn)
		value = x.item("id-tampone")
		call objContentField.insertFieldMatch(6, newMaxID, value, objConn)
		value = x.item("base-tampone")
		call objContentField.insertFieldMatch(7, newMaxID, value, objConn)
		value = x.item("diam-tampone")
		call objContentField.insertFieldMatch(8, newMaxID, value, objConn)
		value = x.item("dim-gomma-A")
		call objContentField.insertFieldMatch(9, newMaxID, value, objConn)
		value = x.item("dim-gomma-B")
		call objContentField.insertFieldMatch(10, newMaxID, value, objConn)
		value = x.item("dim-gomma-C")
		call objContentField.insertFieldMatch(11, newMaxID, value, objConn)
		value = x.item("dim-gomma-D")
		call objContentField.insertFieldMatch(12, newMaxID, value, objConn)
		value = x.item("dim-gomma-E")
		call objContentField.insertFieldMatch(13, newMaxID, value, objConn)
		value = x.item("dim-gomma-F")
		call objContentField.insertFieldMatch(14, newMaxID, value, objConn)
					

		'/**
		'* inserisco i nuovi file allegati
		'*/
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		uploadsDirVar = Application("baseroot")&Application("dir_upload_news")&newMaxID&"/"
		uploadsDirVar = Server.MapPath(uploadsDirVar)		
		call objLogger.write("uploadsDirVar: "&uploadsDirVar, "system", "debug")
		
		uploadsDirVar = uploadsDirVar
		if not(objFSO.FolderExists(uploadsDirVar)) then
			call objFSO.CreateFolder(uploadsDirVar)	
		end if

		filesname = x.item("path-small-img")
		filesname = Right(filesname,(Len(filesname)-InStrRev(filesname,"/",-1,1)))
		filemname = x.item("path-medium-img")	
		filemname = Right(filemname,(Len(filemname)-InStrRev(filemname,"/",-1,1)))			
		call objLogger.write("filesname: "&filesname, "system", "debug")		
		call objLogger.write("filemname: "&filemname, "system", "debug")
		
		fileToUploadS = Application("baseroot")&Application("dir_editor_upload")&"tamponi/"&filesname
		fileToUploadM = Application("baseroot")&Application("dir_editor_upload")&"tamponi/"&filemname
		fileToUploadS = Server.MapPath(fileToUploadS)
		fileToUploadM = Server.MapPath(fileToUploadM)		
		call objLogger.write("fileToUploadS: "&fileToUploadS, "system", "debug")		
		call objLogger.write("fileToUploadM: "&fileToUploadM, "system", "debug")

		if objFSO.FileExists(fileToUploadS)=true then
			call objFSO.CopyFile (fileToUploadS,uploadsDirVar&"\")
		else
		  call objLogger.write("fileToUploadS NOT EXIST", "system", "debug")
		end if

		if objFSO.FileExists(fileToUploadM)=true then
			call objFSO.CopyFile (fileToUploadM,uploadsDirVar&"\")
		else
		  call objLogger.write("fileToUploadM NOT EXIST", "system", "debug")
		end if
	
		Set objFSO = nothing
		
		Set fileXnews = new File4NewsClass	
		tmpFilePath = newMaxID & "/" & filesname					
		new_id_file = fileXnews.insertFile(filesname, "image/jpeg", tmpFilePath, null, 1, objConn)	
		call objLogger.write("tmpFilePath: "&tmpFilePath&" -new_id_file:"&new_id_file, "system", "debug")
		call fileXnews.insertFileXNews(newMaxID, new_id_file, objConn)	
		
		tmpFilePath = newMaxID & "/" & filemname					
		new_id_file = fileXnews.insertFile(filemname, "image/jpeg", tmpFilePath, null, 6, objConn)
		call objLogger.write("tmpFilePath: "&tmpFilePath&" -new_id_file:"&new_id_file, "system", "debug")
		call fileXnews.insertFileXNews(newMaxID, new_id_file, objConn)	
		Set fileXnews = nothing

		'/**
		'* inserisco i nuovi target per news
		'*/
		call objNews.insertTargetXNews(45, newMaxID, objConn)	
		if(x.item("lang-tampone")="1")then
			call objNews.insertTargetXNews(1, newMaxID, objConn)	
		else
			call objNews.insertTargetXNews(3, newMaxID, objConn)	
		end if
		
		'/**
		'* inserisco l'utente per news
		'*/		
		Dim UtenteXnews
		Set UtenteXnews = new UserClass
		call UtenteXnews.insertUserXNews(Session("objCMSUtenteLogged"), newMaxID, objConn)
		Set UtenteXnews = nothing
	next
					
	'if objConn.Errors.Count = 0 then
	'	objConn.CommitTrans
		
	'	'rimuovo gli oggetti find dalla cache
	'	Set objCacheClass = new CacheClass
	'	call objCacheClass.removeByPrefix("findc", null)
	'	Set objCacheClass = nothing	
	'else
		objConn.RollBackTrans
	'	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	'end if
	
	Set objListTamponi = nothing
	
	Set objDB = nothing			
	Set objNews = nothing
	Set objUserLogged = nothing
	response.Redirect(redirectPage)				

	Set objLogger = nothing

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>