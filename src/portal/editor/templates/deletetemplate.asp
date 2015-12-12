<%@ Language=VBScript %>
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
	Dim id_template	
	
	id_template = request("id_template")
	Dim objTempl, objFSO, templateDirVar, objTmpTempl
	Set objTempl = New TemplateClass
	Set objTmpTempl = objTempl.findTemplateByID(id_template)
	templateDirVar = objTmpTempl.getDirTemplate()
	templateDirVar = Application("baseroot") & Application("dir_upload_templ")& templateDirVar
	templateDirVar = Server.MapPath(templateDirVar)
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	if objFSO.FolderExists(templateDirVar) then
		call objFSO.DeleteFolder(templateDirVar, true)	
	end if
	Set objFSO = nothing
	call objTempl.deleteTemplate(id_template)
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