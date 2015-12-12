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

Dim id_template, numBaseTemplate,numMaxImgs, numMaxFiles, numMaxIncludes, numMaxJs
Dim dir_new_template, descrizione_template, fileupload_css_filename, base_template, order_by, elem_x_page
id_template = request("id_template")
dir_new_template = ""
descrizione_template = ""
fileupload_css_filename = ""
base_template = 0
order_by = 1
elem_x_page = 5
numMaxImgs = Application("num_max_attachments")
numMaxFiles = 1
numMaxIncludes = 1
numMaxJs = 1

if(not(request("numMaxImgs") = "")) then
	numMaxImgs = request("numMaxImgs")
end if

if(not(request("numMaxFiles") = "")) then
	numMaxFiles = request("numMaxFiles")
end if

if(not(request("numMaxIncludes") = "")) then
	numMaxIncludes = request("numMaxIncludes")
end if

if(not(request("numMaxJs") = "")) then
	numMaxJs = request("numMaxJs")
end if

if(CInt(id_template) <> -1) then
	Dim objTemp, objCurrTemp
	Set objTemp = New TemplateClass
	Set objCurrTemp = objTemp.findTemplateByID(id_template)
	dir_new_template = objCurrTemp.getDirTemplate()
	descrizione_template = objCurrTemp.getDescrizioneTemplate()
	fileupload_css_filename = objCurrTemp.getTemplateCss()
	base_template = objCurrTemp.getBaseTemplate()
	order_by= objCurrTemp.getOrderBy()
	elem_x_page = objCurrTemp.getElemXPage()
	Set objCurrTemp = nothing
	Set objTemp = nothing
end if
%>