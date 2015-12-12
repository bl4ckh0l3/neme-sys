<%
if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
Set objUserLoggedTmp = nothing
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing

'/**
'* recupero i valori della news selezionata se id_template <> -1
'*/
Dim id_template, dirTemplate, fileCss, descTemplate, baseTemplate, objPages, order_by, elem_x_page
id_template = request("id_template")
dirTemplate = ""
fileCss = ""
descTemplate = ""
baseTemplate = 0
order_by = ""
elem_x_page = ""
objPages = null

if not (isNull(id_template)) then
	Dim objTemplate, objSelTemplate
	Set objTemplate = New TemplateClass
	Set objSelTemplate = objTemplate.findTemplateByID(id_template)
	Set objTemplate = nothing
	
	id_template = objSelTemplate.getID()
	dirTemplate = objSelTemplate.getDirTemplate()
	fileCss = objSelTemplate.getTemplateCss()
	descTemplate = objSelTemplate.getDescrizioneTemplate()
	baseTemplate = objSelTemplate.getBaseTemplate()
	order_by= objSelTemplate.getOrderBy()
	elem_x_page = objSelTemplate.getElemXPage()
	
	On Error Resume Next	
	if not(isNull(objSelTemplate.getPagePerTemplate())) then
		Set objPages = objSelTemplate.getPagePerTemplate()	
	end if	
	if(Err.number <> 0)then
	end if
	Set objSelTemplate = nothing
else
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")			
end if
%>