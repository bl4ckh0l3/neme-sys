<%@ Language=VBScript %>
<% 
'option explicit
'On error resume next 
%>
<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/common/include/Objects/CommentsClass.asp" -->
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
	Dim id_news, dtData_ins, dtData_pub, dtData_del, redirectPage
	Dim objDB, objConn
	redirectPage = Application("baseroot")&"/editor/contenuti/InserisciNews.asp?cssClass=LN&id_news="
	
	Dim objFSO, uploadsDirVar, origUploadsDirVar
	Dim objFileXnews, fileXnews, tmpFileXnews, tmpPath
	Dim xFiles, yFiles
	Dim tmpFileName, tmpFilePath, new_id_file	
	Dim targetXnews	
	Dim xTarget
	Dim DD, MM, YY, HH, MIN, SS
					
	Dim objNews, objOriginalNews
	Set objNews = New NewsClass

	Dim objLocal
	Set objLocal = new LocalizationClass
	
	Dim objLogger
	Set objLogger = New LogClass	

	Set objFileXnews = new File4NewsClass
	
	Dim commentsXnews, commentListXnews
	Set commentsXnews = New CommentsClass

	'/**
	'* recupero la news originale da clonare
	'*/
	On error resume next
	id_news = request("id_news")
	Set objOriginalNews = objNews.findNewsByID(id_news)
	Set targetXnews = objNews.getTargetPerNews(id_news)
	Set fileXnews = objFileXnews.getFilePerNews(id_news)
	Set commentListXnews = commentsXnews.findCommentiByIDElement(id_news, 1, null)
	if(Err.number <> 0) then
		'response.write(Err.description)	
	end if

	On error resume next
	Set points = objLocal.findPointByElement(id_news, 1)
	if(Err.number <> 0) then	
	end if	

	'/**
	'* news da inserire e recupero Max(ID) 
	'*/	
	Dim newMaxID
	dtData_ins = Now()
	dtData_pub = Now()
	dtData_del = ""
	
	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()
	objConn.BeginTrans		
		
	newMaxID = objNews.insertNews(objOriginalNews.getTitolo(), objOriginalNews.getAbstract1(), objOriginalNews.getAbstract2(), objOriginalNews.getAbstract3(), objOriginalNews.getTesto(), objOriginalNews.getKeyword(), dtData_ins, dtData_pub, dtData_del, 0, objOriginalNews.getMetaDescription(), objOriginalNews.getMetaKeyword(), objOriginalNews.getPageTitle(), objConn)
	
	call objLogger.write("clonato contenuto --> titolo: "&objOriginalNews.getTitolo(), objUserLogged.getUserName(), "info")	


	'********** RECUPERO LA LISTA DI FIELD PRODOTTI ASSOCIATI AL CONTENUTO
	Dim objContentField, objListContentField
	Set objContentField = new ContentFieldClass
	On Error Resume Next
	Set objListContentField = objContentField.getListContentField4ContentActive(id_news)

	if(Instr(1, typename(objListContentField), "Dictionary", 1) > 0) then
		if(objListContentField.count > 0)then		
			'/**
			'* inserisco i field per contenuto
			'*/					
			for each xField in objListContentField
				Set objContentField = objListContentField(xField)
				call objContentField.insertFieldMatch(xField, newMaxID, objContentField.getSelValue(), objConn)	
				Set objContentField = nothing
			next
		end if
	end if
	if(Err.number <> 0) then
		call objLogger.write("Errore clonazione field contenuto --> descrizione: "&Err.description, "system", "error")
	end if
	Set objContentField = nothing

	Set objOriginalNews = nothing
	
	'/**
	'* inserisco i nuovi file allegati
	'*/
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	uploadsDirVar = Application("baseroot")&Application("dir_upload_news")
	uploadsDirVar = Server.MapPath(uploadsDirVar)
	origUploadsDirVar = uploadsDirVar
	
	uploadsDirVar = uploadsDirVar & "\" & newMaxID
	if not(objFSO.FolderExists(uploadsDirVar)) then
		call objFSO.CreateFolder(uploadsDirVar)		
	end if

	if(Instr(1, typename(fileXnews), "Dictionary", 1) > 0) then
		On error resume next
		for each xFiles in fileXnews	
			Set tmpFileXnews = fileXnews(xFiles)
			tmpFileName = tmpFileXnews.getFileName()
			tmpFilePath = newMaxID & "/" & tmpFileName					
			new_id_file = objFileXnews.insertFile(tmpFileName, tmpFileXnews.getFileType(), tmpFilePath, tmpFileXnews.getFileDida(), tmpFileXnews.getFileTypeLabel(), objConn)
			call objFileXnews.insertFileXNews(newMaxID, new_id_file, objConn)	
			Set tmpFileXnews = nothing
		next	
		Set fileXnews = nothing	
		if (objFSO.FolderExists(objFSO.BuildPath(origUploadsDirVar,id_news))) then
			objFSO.CopyFile objFSO.BuildPath(origUploadsDirVar,id_news)&"\*.*",uploadsDirVar&"\"
		end if
		if(Err.number <> 0) then	
		end if
	end if

	Set objFileXnews = nothing
	Set objFSO = nothing

	'/**
	'* inserisco i target per news clonati
	'*/
	On error resume next
	for each xTarget in targetXnews
		call objNews.insertTargetXNews(xTarget, newMaxID, objConn)	
	next			
	Set targetXnews = nothing
	if(Err.number <> 0) then	
	end if	
	
	'/**
	'* inserisco l'utente per news
	'*/	
	On error resume next	
	Dim UtenteXnews
	Set UtenteXnews = new UserClass
	call UtenteXnews.insertUserXNews(Session("objCMSUtenteLogged"), newMaxID, objConn)
	Set UtenteXnews = nothing
	if(Err.number <> 0) then	
	end if
	
	'/**
	'* inserisco i commenti per news
	'*/	
	On error resume next
	for each xComment in commentListXnews.Items
		call commentsXnews.insertCommento(newMaxID, xComment.getElementType(), xComment.getIDUtente(), xComment.getMessage(), xComment.getVoteType(), xComment.getActive(), objConn)
	next	
	Set commentListXnews = nothing
	Set commentsXnews = nothing
	if(Err.number <> 0) then
	end if	

	'/**
	'* inserisco i dati di geolocalizzazione clonati
	'*/					
	if not(isNull(points)) then
		'call objLogger.Write("insertPoint --> newMaxID: "&newMaxID&" - xLocal.getLatitude():"&xLocal.getLatitude()&"xLocal.getLongitude():"&xLocal.getLongitude()&"xLocal.getInfo():"&xLocal.getInfo(), objUserLogged.getUserName(), "info")
		for each xLocal in points.Items
			call objLocal.insertPoint(newMaxID, 1, xLocal.getLatitude(), xLocal.getLongitude(), xLocal.getInfo(), objConn)
		next
	end if	
	Set objLocal = nothing
					
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
	Set objNews = nothing
	Set objUserLogged = nothing	

	response.Redirect(redirectPage&newMaxID)	

	Set objLogger = nothing

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>