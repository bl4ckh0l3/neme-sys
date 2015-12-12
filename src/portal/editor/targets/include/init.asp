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

Dim objTarget, objListaTarget
Set objTarget = New TargetClass

Dim totPages, itemsXpage, numPage

if not(request("items") = "") then
	session("targetItems") = request("items")
	itemsXpage = session("targetItems")
	session("targetPage") = 1
else
	if not(session("targetItems") = "") then
		itemsXpage = session("targetItems")
	else
		session("targetItems") = 20
		itemsXpage = session("targetItems")
	end if
end if

if not(request("page") = "") then
	session("targetPage") = request("page")
	numPage = session("targetPage")
else
	if not(session("targetPage") = "") then
		numPage = session("targetPage")
	else
		session("targetPage") = 1
		numPage = session("targetPage")
	end if
end if	
%>