<%@Language=VBScript codepage=65001 %>
<% 
'option explicit
On error resume next 
Server.ScriptTimeout=3600 ' max value = 2147483647
Response.Expires=-1500
Response.Buffer = TRUE
Response.Clear
Response.ContentType="text/html"
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->

<%
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
	
	'/**
	'* Recupero tutti i parametri dal form e li elaboro
	'*/	
	Dim id_news, strTitolo, strAbs, strAbs2, strAbs3, strText, strKeyword, strServerName, dtData_ins
	Dim dtData_pub, dtData_del, stato_news, send_newsletter, reqTargets, reqFiles, arrTarget, arrFiles
	Dim reqFilesToMod, arrFilesToMod, save_esc, redirectPage, numMaxImgs
     Dim Upload, fileName, fileSize, ks, i, fileKey
	Dim uploadsDirVar, choose_newsletter
	Dim page_title, meta_description, meta_keyword
	Dim objDB, objConn
	
     Set Upload = Server.CreateObject("Persits.Upload")
	Upload.CodePage = 65001
	Upload.Save
	
	strServerName = Upload.Form("srv_name")
	redirectPage = Application("baseroot")&"/editor/contenuti/ListaNews.asp"
	save_esc = Upload.Form("save_esc")	
	id_news = Upload.Form("id_news")
	strTitolo = Upload.Form("titolo")
	'** sostituisco dal titolo:
		'èéàòùì
	'** con:
		'&egrave;&eacute;&agrave;&ograve;&ugrave;&igrave;
	'strTitolo = Replace(strTitolo, "'", "&#39;", 1, -1, 1)
	'strTitolo = Replace(strTitolo, "è", "&egrave;", 1, -1, 1)
	'strTitolo = Replace(strTitolo, "é", "&eacute;", 1, -1, 1)
	'strTitolo = Replace(strTitolo, "à", "&agrave;", 1, -1, 1)
	'strTitolo = Replace(strTitolo, "ò", "&ograve;", 1, -1, 1)
	'strTitolo = Replace(strTitolo, "ù", "&ugrave;", 1, -1, 1)
	'strTitolo = Replace(strTitolo, "ì", "&igrave;", 1, -1, 1)
	
	'** sostituisco dai campi abstract e dal testo i caratteri di default aggiunti dall'editor html:
	strAbs = Upload.Form("abstract1")
	if (strAbs ="<br type=&quot;_moz&quot; />" or strAbs ="<br type=""_moz"" />" or strAbs ="&lt;br type=&quot;_moz&quot; /&gt;" or strAbs ="&lt;br /&gt;" or strAbs ="<br />") then
		strAbs = ""
	else
		strAbs = Replace(strAbs, "src="""&strServerName, "src=""", 1, -1, 1)
		strAbs = Replace(strAbs, "\r\n", "", 1, -1, 1)
		strAbs = Replace(strAbs, "'", "&#39;", 1, -1, 1)
	end if
	
	strAbs2 = Upload.Form("abstract2")
	if (strAbs2 ="<br type=&quot;_moz&quot; />" or strAbs2 ="<br type=""_moz"" />" or strAbs2 ="&lt;br type=&quot;_moz&quot; /&gt;" or strAbs2 ="&lt;br /&gt;" or strAbs2 ="<br />") then
		strAbs2 = ""
	else
		strAbs2 = Replace(strAbs2, "src="""&strServerName, "src=""", 1, -1, 1)
		strAbs2 = Replace(strAbs2, "\r\n", "", 1, -1, 1)
		strAbs2 = Replace(strAbs2, "'", "&#39;", 1, -1, 1)
	end if
	
	strAbs3 = Upload.Form("abstract3")
	if (strAbs3 ="<br type=&quot;_moz&quot; />" or strAbs3 ="<br type=""_moz"" />" or strAbs3 ="&lt;br type=&quot;_moz&quot; /&gt;" or strAbs3 ="&lt;br /&gt;" or strAbs3 ="<br />") then
		strAbs3 = ""
	else
		strAbs3 = Replace(strAbs3, "src="""&strServerName, "src=""", 1, -1, 1)
		strAbs3 = Replace(strAbs3, "\r\n", "", 1, -1, 1)
		strAbs3 = Replace(strAbs3, "'", "&#39;", 1, -1, 1)
	end if
	
	strText = Upload.Form("testo")
	if (strText ="<br type=&quot;_moz&quot; />" or strText ="<br type=""_moz"" />" or strText ="&lt;br type=&quot;_moz&quot; /&gt;" or strText ="&lt;br /&gt;" or strText ="<br />") then
		strText = ""
	else
		strText = Replace(strText, "src="""&strServerName, "src=""", 1, -1, 1)
		strText = Replace(strText, "\r\n", "", 1, -1, 1)
		strText = Replace(strText, "'", "&#39;", 1, -1, 1)	
	end if
	
	strKeyword = Upload.Form("keyword")
	dtData_ins = Upload.Form("news_data")
	dtData_pub = Upload.Form("news_data_pub")
	dtData_del = Upload.Form("news_data_del")
	stato_news = Upload.Form("stato_news")	
	send_newsletter = -1
	if not(Upload.Form("send_newsletter") = "") then
		send_newsletter = Upload.Form("send_newsletter")
	end if
	choose_newsletter = -1
	if not(Upload.Form("choosenNewsletter") = "") then
		choose_newsletter = Upload.Form("choosenNewsletter")
	end if	
	reqTargets = Upload.Form("ListTarget")
	
	if(reqTargets="") then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=006")
	end if
	
	page_title = Upload.Form("page_title")
	meta_description = Upload.Form("meta_description")
	meta_keyword = Upload.Form("meta_keyword")
	
	reqFiles = Upload.Form("ListFileDaEliminare")
	reqFilesToMod = Upload.Form("ListFileDaModificare")
	numMaxImgs = Upload.Form("numMaxImgs")

	reqFieldList = Upload.Form("list_content_fields")
	'reqFieldListValues = Upload.Form("list_content_fields_values")
		
	arrTarget = split(reqTargets, "|", -1, 1)
	arrFiles = split(reqFiles, "|", -1, 1)
	arrFilesToMod = split(reqFilesToMod, "|", -1, 1)
	arrFieldList = split(reqFieldList, "|", -1, 1)
	'arrFieldListValues = split(reqFieldListValues, "##", -1, 1)	
	
	Dim objFSO
	Dim fileXnews, tmpFileXnews, tmpPath
	Dim xFiles, yFiles
	Dim tmpFileName, tmpFilePath, new_id_file	
	Dim targetXnews	
	Dim xTarget
	Dim DD, MM, YY, HH, MIN, SS
	Dim FileUploaded
					
	Dim objNews
	Set objNews = New NewsClass
	
	Dim objLogger
	Set objLogger = New LogClass

	Dim objContentField
	Set objContentField = new ContentFieldClass
			
	Dim objNewsLetter, objNewsletterNews

	if (Cint(id_news) <> -1) then
		'/**
		'* news da mofificare
		'*/		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		objConn.BeginTrans		

		call objNews.modifyNews(id_news, strTitolo, strAbs, strAbs2, strAbs3, strText, strKeyword, dtData_ins, dtData_pub, dtData_del, stato_news, meta_description, meta_keyword, page_title, objConn)		
		call objLogger.write("modificato contenuto --> id: "&id_news&"; titolo: "&strTitolo, objUserLogged.getUserName(), "info")
		
		'/**
		'* cancello i vecchi field e inserisco i nuovi field per contenuto
		'*/			
		call objContentField.deleteFieldMatchByContent(id_news, objConn)
		
		for each xField in arrFieldList
			idField = Left(xField,InStr(1,xField,"-",1)-1)
			value = ""
			typeField = Mid(xField,InStr(1,xField,"-",1)+1)					
			select Case Cint(typeField)
			Case 4,5,6
				value = Upload.Form("hidden_"&objContentField.getFieldPrefix()&idField)
			Case Else
				value = Upload.Form(objContentField.getFieldPrefix()&idField)			
			End Select			
			call objContentField.insertFieldMatch(idField, id_news, value, objConn)	
		next	

		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		uploadsDirVar = Application("baseroot")&Application("dir_upload_news")
		uploadsDirVar = Server.MapPath(uploadsDirVar)

		Set fileXnews = new File4NewsClass
		
		'/**
		'* modifico le didascalie dei file allegati!
		'*/			
		Dim strDida, strFileTypeLabel, objSingleFile
		
		for each yFiles in arrFilesToMod
			Set objSingleFile = fileXnews.getFileByID(yFiles)
			strDida = Upload.Form("fileDaModificare_"&objSingleFile.getFileID())
			strFileTypeLabel = Upload.Form("fileDaModificare_"&objSingleFile.getFileID()&"_label")
			call fileXnews.modifyFile(objSingleFile.getFileID(), objSingleFile.getFileName(), objSingleFile.getFileType(), objSingleFile.getFilePath(), strDida, strFileTypeLabel, objConn)
			Set objSingleFile = nothing
		next		
		
		'/**
		'* cancello i file allegati selezionati
		'*/		
		for each xFiles in arrFiles
			Set tmpFileXnews = fileXnews.getFileByID(xFiles)			
			tmpPath = tmpFileXnews.getFilePath()			
			tmpPath = Replace(tmpPath, "/", "\", 1, -1, 1)
			call fileXnews.deleteFileXNews(xFiles, objConn)			
			if(objFSO.FileExists(uploadsDirVar & "\" & tmpPath)) then
				objFSO.DeleteFile(uploadsDirVar & "\" & tmpPath)
			end if
			call fileXnews.deleteFile(xFiles, objConn)
			Set tmpFileXnews = nothing	
		next
		
		'/**
		'* inserisco i nuovi file allegati
		'*/
		uploadsDirVar = uploadsDirVar & "\" & id_news		
		if not(objFSO.FolderExists(uploadsDirVar)) then
'<!--nsys-demonwsproc1-->
			call objFSO.CreateFolder(uploadsDirVar)	
'<!---nsys-demonwsproc1-->
		end if
		Set objFSO = nothing

		
		ks = Upload.Files.count
		if (ks > 0) then
			dim f
			for f = 1 to numMaxImgs
				if(Instr(1, typename(Upload.Files("fileupload"&f)), "IUploadedFile", 1) > 0) then
					Set FileUploaded = Upload.Files("fileupload"&f)
					tmpFileName = FileUploaded.FileName	
					tmpFilePath = id_news & "/" & tmpFileName
					if (Instr(1, typename(fileXnews.getFileByFileNameAndIdNews(id_news,tmpFileName)), "File4NewsClass", 1) > 0) then
						call fileXnews.modifyFile(fileXnews.getFileByFileNameAndIdNews(id_news,tmpFileName).getFileID(),tmpFileName, FileUploaded.ContentType, tmpFilePath, Upload.Form("fileupload"&f & "_dida"), Upload.Form("fileupload"&f & "_label"), objConn)
					else
						new_id_file = fileXnews.insertFile(tmpFileName, FileUploaded.ContentType, tmpFilePath, Upload.Form("fileupload"&f & "_dida"), Upload.Form("fileupload"&f & "_label"), objConn)
						call fileXnews.insertFileXNews(id_news, new_id_file, objConn)					
					end if
'<!--nsys-demonwsproc2-->					
					FileUploaded.SaveAs(uploadsDirVar & "\" & FileUploaded.Filename)
'<!---nsys-demonwsproc2-->
					Set FileUploaded = nothing
				end if
			next
		end if		
		Set fileXnews = nothing


		'/**
		'* cancello i vecchi target e inserisco i nuovi target per news
		'*/			
		call objNews.deleteTargetXNews(id_news, objConn)
		
		for each xTarget in arrTarget
			call objNews.insertTargetXNews(xTarget, id_news, objConn)	
		next
						
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans

			'rimuovo l'oggetto dalla cache
			Set objCacheClass = new CacheClass
			call objCacheClass.remove("content-"&id_news)
			call objCacheClass.remove("listcf-"&id_news)
			call objCacheClass.removeByPrefix("findc", id_news)
			Set objCacheClass = nothing
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		Set objDB = nothing
						
		'/*
		' * controllo se effettuare l'invio della newsletter
		' */
		if(Cint(send_newsletter) = 1) then		
			Set objNewsLetter = new NewsletterClass
			call objNewsLetter.sendNewsletter(id_news, choose_newsletter)
			Set objNewsLetter = nothing	
		end if
		
		Set objNews = nothing
		Set objUserLogged = nothing
	
		if(save_esc = 0) then
			redirectPage = Application("baseroot")&"/editor/contenuti/inseriscinews.asp?cssClass=LN&id_news="&id_news
		end if
		response.Redirect(redirectPage)		
	else
		'/**
		'* news da inserire e recupero Max(ID) 
		'*/	
		Dim newMaxID
		dtData_ins = Now()
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		objConn.BeginTrans		
			
		newMaxID = objNews.insertNews(strTitolo, strAbs, strAbs2, strAbs3, strText, strKeyword, dtData_ins, dtData_pub, dtData_del, stato_news, meta_description, meta_keyword, page_title, objConn)		
		call objLogger.write("inserito contenuto --> titolo: "&strTitolo, objUserLogged.getUserName(), "info")		
		
		'/**
		'* cancello i vecchi field e inserisco i nuovi field per contenuto
		'*/			
		call objContentField.deleteFieldMatchByContent(newMaxID, objConn)
		
		for each xField in arrFieldList
			idField = Left(xField,InStr(1,xField,"-",1)-1)
			value = ""
			typeField = Mid(xField,InStr(1,xField,"-",1)+1)					
			select Case Cint(typeField)
			Case 4,5,6
				value = Upload.Form("hidden_"&objContentField.getFieldPrefix()&idField)
			Case Else
				value = Upload.Form(objContentField.getFieldPrefix()&idField)			
			End Select			
			call objContentField.insertFieldMatch(idField, newMaxID, value, objConn)	
		next	
		
		'/**
		'* inserisco i nuovi file allegati
		'*/
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		uploadsDirVar = Application("baseroot")&Application("dir_upload_news")
		uploadsDirVar = Server.MapPath(uploadsDirVar)
		
		uploadsDirVar = uploadsDirVar & "\" & newMaxID
		if not(objFSO.FolderExists(uploadsDirVar)) then
'<!--nsys-demonwsproc3-->
			call objFSO.CreateFolder(uploadsDirVar)	
'<!---nsys-demonwsproc3-->	
		end if
		Set objFSO = nothing
		
		Set fileXnews = new File4NewsClass
		
		ks = Upload.Files.count
		if (ks > 0) then
			dim q
			for q = 1 to numMaxImgs
				if(Instr(1, typename(Upload.Files("fileupload"&q)), "IUploadedFile", 1) > 0) then
					Set FileUploaded = Upload.Files("fileupload"&q)
					tmpFileName = FileUploaded.FileName
					tmpFilePath = newMaxID & "/" & tmpFileName					
					new_id_file = fileXnews.insertFile(tmpFileName, FileUploaded.ContentType, tmpFilePath, Upload.Form("fileupload"&q & "_dida"), Upload.Form("fileupload"&q & "_label"), objConn)
					call fileXnews.insertFileXNews(newMaxID, new_id_file, objConn)
'<!--nsys-demonwsproc4-->
					FileUploaded.SaveAs(uploadsDirVar & "\" & FileUploaded.Filename)
'<!---nsys-demonwsproc4-->
					Set FileUploaded = nothing
				end if
			next
		end if		
		Set fileXnews = nothing

		'/**
		'* cancello i vecchi target e inserisco i nuovi target per news
		'*/
		for each xTarget in arrTarget
			call objNews.insertTargetXNews(xTarget, newMaxID, objConn)	
		next				
		
		'/**
		'* inserisco l'utente per news
		'*/		
		Dim UtenteXnews
		Set UtenteXnews = new UserClass
		call UtenteXnews.insertUserXNews(Session("objCMSUtenteLogged"), newMaxID, objConn)
		Set UtenteXnews = nothing

		'/**
		'* aggiorno le localizzazioni se sono state inserite prima di salvare il contenuto
		'*/
		if(Upload.Form("pregeoloc_el_id")<>"") then
			Set objLoc = new LocalizationClass
			Set listOfPoints = objLoc.findPointByElement(Upload.Form("pregeoloc_el_id"), 1)
			for each q in listOfPoints
				call objLoc.modifyPoint(q, newMaxID, listOfPoints(q).getLatitude(), listOfPoints(q).getLongitude(), listOfPoints(q).getInfo(), objConn)
			next
			Set listOfPoints = nothing
			Set objLoc = nothing
		end if
						
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
			
			'rimuovo gli oggetti find dalla cache
			Set objCacheClass = new CacheClass
			call objCacheClass.removeByPrefix("findc", null)
			Set objCacheClass = nothing	
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		Set objDB = nothing
		
		'/*
		' * controllo se effettuare l'invio della newsletter
		' */
		if(Cint(send_newsletter) = 1) then
			Set objNewsLetter = new NewsletterClass
			call objNewsLetter.sendNewsletter(newMaxID, choose_newsletter)
			Set objNewsLetter = nothing	
		end if
				
		Set objNews = nothing
		Set objUserLogged = nothing
		
		if(save_esc = 0) then
			redirectPage = Application("baseroot")&"/editor/contenuti/inseriscinews.asp?cssClass=LN&id_news="&newMaxID
		end if
		response.Redirect(redirectPage)				
	end if

	Set objLogger = nothing

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>