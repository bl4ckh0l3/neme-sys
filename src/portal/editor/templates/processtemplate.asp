<%@ Language=VBScript %>
<% 
On error resume next 
Response.Expires = -1
Server.ScriptTimeout = 600
%>
<!-- #include virtual="/common/include/Objects/FileUploadClass.asp" -->
<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->

<%
if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if
	
	'/**
	'* Recupero tutti i parametri dal form e li elaboro
	'*/	
    Dim Upload
	Dim uploadsDirVar, uploadsDirVarAsp, uploadsDirVarCss, uploadsDirVarImgs, uploadsDirVarIncludes, uploadsDirVarJs, uploadsDirVarAspLang
	Dim id_template, dir_new_template, fileupload_name, descrizione_template, base_template, fileupload_css, order_by, elem_x_page
	Dim numMaxFilesToUpload, numMaxIncludes
	
    Set Upload = New FileUploadClass
	Upload.SaveField()
	
	id_template = Upload.Form("id_template")
	
	if not(id_template = "") then
		dir_new_template = Upload.Form("dir_new_template")
		descrizione_template = Upload.Form("descrizione_template")
		base_template = Upload.Form("base_template")
		base_template = CInt(base_template)	
		numMaxFilesToUpload =  Upload.Form("numMaxFilesToUpload")
		numMaxIncludes = Upload.Form("numMaxIncludes")
		fileupload_css = Upload.Form("fileupload_css_filename")		
		order_by = Upload.Form("order_by")
		elem_x_page = Upload.Form("elem_x_page")
						
		Dim objTemp, objFSO, objPagePerTemp
		Set objTemp = New TemplateClass
		Set objPagePerTemp = new Page4TemplateClass
		
		uploadsDirVar = Application("baseroot")&Application("dir_upload_templ")	
		uploadsDirVar = Server.MapPath(uploadsDirVar)	
		uploadsDirVarAsp = uploadsDirVar & "\" & dir_new_template &"\"
		uploadsDirVarCss = uploadsDirVarAsp &"css\"
		uploadsDirVarImgs = uploadsDirVarAsp &"img\"
		uploadsDirVarIncludes =  uploadsDirVarAsp &"include\"
		uploadsDirVarJs =  uploadsDirVarAsp &"js\"
		
		'*** RECUPERO LA LISTA DI LANGUAGE DISPONIBILI E LA PASSO
		'*** ALLA CLASSE DI UPLOAD, PER SALVARE TUTTI I FILE ASP
		'*** PRINCIPALI
		Dim objSelLanguage, objLangList
		Set objSelLanguage = New LanguageClass
		Set objLangList = objSelLanguage.getListaLanguage()
		Set objSelLanguage = nothing
		
		'response.Write("id_template: " & id_template & "<br>")
		'response.Write("dir_new_template: " & dir_new_template & "<br>")
		'response.Write("descrizione_template: " & descrizione_template & "<br>")
		'response.Write("base_template: " & base_template & "<br>")
		'response.Write("numMaxFilesToUpload: " & numMaxFilesToUpload & "<br>")
		'response.Write("fileupload_css: " & fileupload_css & "<br>")
		'response.Write("uploadsDirVar: " & uploadsDirVar & "<br>")
		'response.Write("uploadsDirVarAsp: " & uploadsDirVarAsp & "<br>")
		'response.Write("uploadsDirVarCss: " & uploadsDirVarCss & "<br>")
		'response.Write("uploadsDirVarImgs: " & uploadsDirVarImgs & "<br>")
		'response.Write("uploadsDirVarIncludes: " & uploadsDirVarIncludes & "<br>")
		'response.Write("uploadsDirVarJs: " & uploadsDirVarJs & "<br>")
	
		'/**
		'* inserisco i nuovi file allegati
		'*/
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		if not(objFSO.FolderExists(uploadsDirVarAsp)) then
			objFSO.CreateFolder(uploadsDirVarAsp)	
		end if
		if not(objFSO.FolderExists(uploadsDirVarCss)) then
			objFSO.CreateFolder(uploadsDirVarCss)	
		end if
		if not(objFSO.FolderExists(uploadsDirVarImgs)) then
			objFSO.CreateFolder(uploadsDirVarImgs)	
		end if
		if not(objFSO.FolderExists(uploadsDirVarIncludes)) then
			objFSO.CreateFolder(uploadsDirVarIncludes)	
		end if
		if not(objFSO.FolderExists(uploadsDirVarJs)) then
			objFSO.CreateFolder(uploadsDirVarJs)	
		end if
		
		For Each lang In objLangList
			uploadsDirVarAspLang = uploadsDirVarAsp & Ucase(objLangList(lang).getLanguageDescrizione()) & "\"
			uploadsDirVarAspLangInclude = uploadsDirVarAspLang &"include\"
			if not(objFSO.FolderExists(uploadsDirVarAspLang)) then
				objFSO.CreateFolder(uploadsDirVarAspLang)	
				objFSO.CreateFolder(uploadsDirVarAspLangInclude)
			end if
		Next
		
		call Upload.SaveTemplatesFiles(objLangList, uploadsDirVarAsp, uploadsDirVarCss, uploadsDirVarImgs, uploadsDirVarIncludes, uploadsDirVarJs)

		Set objLangList = nothing	
		
	
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		objConn.BeginTrans			
	
		'/**
		'* template da inserire
		'*/	
		if(Cint(id_template) = -1) then
			id_template = objTemp.insertTemplate(dir_new_template, fileupload_css, descrizione_template, base_template, order_by, elem_x_page, objConn)		
		else
			call objTemp.modifyTemplate(id_template, dir_new_template, fileupload_css, descrizione_template, base_template, order_by, elem_x_page, objConn)
		end if
		
		Dim filename, position	
		for y = 1 To numMaxFilesToUpload		
			filename = Upload.Form("fileupload_filename_send_"&y)
			position = Upload.Form("fileupload_position_"&y)
			if not(filename="") AND not(position="") then
				if not(isNull(objPagePerTemp.findPagePerTemplateByFileName(id_template, filename)))then
					call objPagePerTemp.modifyPagePerTemplate(objPagePerTemp.findPagePerTemplateByFileName(id_template, filename).getID(),id_template, filename, position, objConn)
				else
					call objPagePerTemp.insertPagePerTemplate(id_template, filename, position, objConn)				
				end if
			end if
		next	
		
		for z = 1 To numMaxIncludes		
			filename = Upload.Form("fileupload_include_send_"&z)
			if not(filename="") then
				if not(isNull(objPagePerTemp.findPagePerTemplateByFileName(id_template, filename)))then
					call objPagePerTemp.modifyPagePerTemplate(objPagePerTemp.findPagePerTemplateByFileName(id_template, filename).getID(),id_template, filename, -1, objConn)
				else
					call objPagePerTemp.insertPagePerTemplate(id_template, filename, -1, objConn)				
				end if
			end if
		next
		
			
		if objConn.Errors.Count = 0 AND Err.Number = 0 then
			objConn.CommitTrans
		else			
			if (objFSO.FolderExists(uploadsDirVarAsp)) then
				objFSO.DeleteFolder(uploadsDirVarAsp)	
			end if
			if (objFSO.FolderExists(uploadsDirVarCss)) then
				objFSO.DeleteFolder(uploadsDirVarCss)	
			end if
			if (objFSO.FolderExists(uploadsDirVarImgs)) then
				objFSO.DeleteFolder(uploadsDirVarImgs)	
			end if
			if (objFSO.FolderExists(uploadsDirVarIncludes)) then
				objFSO.DeleteFolder(uploadsDirVarIncludes)	
			end if
			if (objFSO.FolderExists(uploadsDirVarJs)) then
				objFSO.DeleteFolder(uploadsDirVarJs)	
			end if
			objConn.RollBackTrans	
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
		Set objDB = Nothing	
		
		Set objTemp = nothing
		Set objPagePerTemp = nothing
		Set objUserLogged = nothing
		response.Redirect(Application("baseroot")&"/editor/templates/ListaTemplates.asp")				
	
		' If something fails inside the script, but the exception is handled
		If Err.Number<>0 then
			if (objFSO.FolderExists(uploadsDirVarAsp)) then
				objFSO.DeleteFolder(uploadsDirVarAsp)	
			end if
			if (objFSO.FolderExists(uploadsDirVarCss)) then
				objFSO.DeleteFolder(uploadsDirVarCss)	
			end if
			if (objFSO.FolderExists(uploadsDirVarImgs)) then
				objFSO.DeleteFolder(uploadsDirVarImgs)	
			end if
			if (objFSO.FolderExists(uploadsDirVarIncludes)) then
				objFSO.DeleteFolder(uploadsDirVarIncludes)	
			end if
			if (objFSO.FolderExists(uploadsDirVarJs)) then
				objFSO.DeleteFolder(uploadsDirVarJs)	
			end if			
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		Set objFSO = nothing
	else
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=030")
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>