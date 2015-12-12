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

Dim objTemplates, objListaTemplates
Set objTemplates = New TemplateClass

Dim order_templ_by, reqTemplBy
order_templ_by = null
reqTemplBy = request("order_by")

if (not(isNull(reqTemplBy)) AND not(reqTemplBy = "")) then
	order_templ_by = reqTemplBy	
end if

Dim totPages, itemsXpage, numPage

if not(request("items") = "") then
	session("templateItems") = request("items")
	itemsXpage = session("templateItems")
	session("templatePage") = 1
else
	if not(session("templateItems") = "") then
		itemsXpage = session("templateItems")
	else
		session("templateItems") = 20
		itemsXpage = session("templateItems")
	end if
end if

if not(request("page") = "") then
	session("templatePage") = request("page")
	numPage = session("templatePage")
else
	if not(session("templatePage") = "") then
		numPage = session("templatePage")
	else
		session("templatePage") = 1
		numPage = session("templatePage")
	end if
end if	
%>