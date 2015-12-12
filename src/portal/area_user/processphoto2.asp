<%@ Language=VBScript %>
<% 
option explicit
On error resume next 
Response.Expires = -1
Server.ScriptTimeout = 1200
%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/UserFilesClass.asp" -->


<%
if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	
	'/**
	'* Recupero tutti i parametri dal form e li elaboro
	'*/	
	Dim redirectPage, numMaxImgs
     Dim Upload, fileName, fileSize, ks, i, fileKey
	Dim uploadsDirVar
	Dim objDB, objConn
	
     Set Upload = Server.CreateObject("Persits.Upload")
	Upload.Save
	
	redirectPage = Application("baseroot")&"/area_user/userphotos.asp"

	numMaxImgs = Upload.Form("numMaxImgs")
	
	Dim objFSO, FileUploaded
	Dim fileXnews, tmpFileXnews, tmpPath
	Dim xFiles, yFiles
	Dim tmpFileName, tmpFilePath, new_id_file, tmpFileDida	
					
	Dim objFiles
	Set objFiles = New UserFilesClass
	
	Dim objLogger
	Set objLogger = New LogClass			
		
	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()
	objConn.BeginTrans				
	
	'/**
	'* inserisco i nuovi file allegati
	'*/
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	uploadsDirVar = Application("baseroot")&Application("dir_upload_user")
	uploadsDirVar = Server.MapPath(uploadsDirVar)		
	uploadsDirVar = uploadsDirVar & "\" & objUserLogged.getUserID()
	
	if not(objFSO.FolderExists(uploadsDirVar)) then
		call objFSO.CreateFolder(uploadsDirVar)		
	end if
	Set objFSO = nothing
		
	ks = Upload.Files.count
	if (ks > 0) then
		dim q
		for q = 1 to numMaxImgs
			if(Instr(1, typename(Upload.Files("fileupload"&q)), "IUploadedFile", 1) > 0) then
				Set FileUploaded = Upload.Files("fileupload"&q)
				tmpFileName = FileUploaded.FileName
				tmpFilePath = objUserLogged.getUserID() & "/" & tmpFileName		
				tmpFileDida = Upload.Form("fileupload"&q & "_dida")
				tmpFileDida = Replace(tmpFileDida, "'", "&#39;", 1, -1, 1)
				tmpFileDida = Replace(tmpFileDida, "è", "&egrave;", 1, -1, 1)
				tmpFileDida = Replace(tmpFileDida, "é", "&eacute;", 1, -1, 1)
				tmpFileDida = Replace(tmpFileDida, "à", "&agrave;", 1, -1, 1)
				tmpFileDida = Replace(tmpFileDida, "ò", "&ograve;", 1, -1, 1)
				tmpFileDida = Replace(tmpFileDida, "ù", "&ugrave;", 1, -1, 1)
				tmpFileDida = Replace(tmpFileDida, "ì", "&igrave;", 1, -1, 1)
				
				new_id_file = objFiles.insertFiles(objUserLogged.getUserID(), tmpFileName, FileUploaded.ContentType, tmpFilePath, tmpFileDida, Upload.Form("fileupload"&q & "_label"), objConn)	
				'call objFiles.insertFilesNoTransaction(objUserLogged.getUserID(), tmpFileName, FileUploaded.ContentType(), tmpFilePath, tmpFileDida, Upload.Form("fileupload"&q & "_label"))	
	
				call objLogger.write("inserito file per utente --> filename: "&tmpFileName, objUserLogged.getUserName(), "info")	

				FileUploaded.SaveAs(uploadsDirVar & "\" & FileUploaded.Filename)
				Set FileUploaded = nothing
			end if
		next
	end if		
					
	if objConn.Errors.Count = 0 then
		objConn.CommitTrans
	else
		objConn.RollBackTrans
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
	
	Set objDB = nothing
			
	Set objFiles = nothing
	Set objUserLogged = nothing	
	Set objLogger = nothing

	response.Redirect(redirectPage)	

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>