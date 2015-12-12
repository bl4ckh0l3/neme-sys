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
	Dim id_template, new_dir_template, id_new_template	
	
	id_template = request("id_template")
	new_dir_template = request("new_dir_template")
	Dim objTempl, objFSO, oldTemplateDirVar, newTemplateDirVar, objTmpTempl
	Set objTempl = New TemplateClass
	Set objPagePerTemp = new Page4TemplateClass
	Set objTmpTempl = objTempl.findTemplateByID(id_template)
	oldTemplateDirVar = objTmpTempl.getDirTemplate()
	oldTemplateDirVar = Application("baseroot") & Application("dir_upload_templ")& oldTemplateDirVar
	oldTemplateDirVar = Server.MapPath(oldTemplateDirVar)&"\*"
	newTemplateDirVar = Server.MapPath(Application("baseroot") & Application("dir_upload_templ"))& "\" & new_dir_template &"\"
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	if not(objFSO.FolderExists(oldTemplateDirVar)) then
		objFSO.CreateFolder(newTemplateDirVar)	
		objFSO.CopyFile oldTemplateDirVar, newTemplateDirVar, false
		objFSO.CopyFolder oldTemplateDirVar, newTemplateDirVar, false	
		
		' *** aggiorno il DB con il nuovo template clonato		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		objConn.BeginTrans		
		
		id_new_template = objTmpTempl.insertTemplate(new_dir_template, objTmpTempl.getTemplateCss(), new_dir_template, Cint(objTmpTempl.getBaseTemplate()), Cint(objTmpTempl.getOrderBy()), Cint(objTmpTempl.getElemXPage()), objConn)
		
		Set objTemplPages = objTmpTempl.getPagePerTemplate()
		for each j in objTemplPages
			call objPagePerTemp.insertPagePerTemplate(id_new_template, objTemplPages(j).getFileName(), objTemplPages(j).getPageNum(), objConn)
		next		
		Set objTemplPages = nothing

		if objConn.Errors.Count = 0 AND Err.Number = 0 then
			objConn.CommitTrans
		else			
			if (objFSO.FolderExists(newTemplateDirVar)) then
				objFSO.DeleteFolder(newTemplateDirVar)	
			end if
			objConn.RollBackTrans	
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
		Set objDB = Nothing

	end if
	Set objFSO = nothing
	
	
	Set objPagePerTemp = nothing
	Set objTmpTempl = nothing
	Set objTempl = nothing
	
	Set objUserLogged = nothing
	response.Redirect(Application("baseroot")&"/editor/templates/ListaTemplates.asp")				


	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>