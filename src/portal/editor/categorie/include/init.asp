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

Dim objCategoria, objListaCategorie
Set objCategoria = New CategoryClass

Dim totPages, itemsXpage, numPage

if not(request("items") = "") then
	session("categorieItems") = request("items")
	itemsXpage = session("categorieItems")
	session("categoriePage") = 1
else
	if not(session("categorieItems") = "") then
		itemsXpage = session("categorieItems")
	else
		session("categorieItems") = 20
		itemsXpage = session("categorieItems")
	end if
end if

if not(request("page") = "") then
	session("categoriePage") = request("page")
	numPage = session("categoriePage")
else
	if not(session("categoriePage") = "") then
		numPage = session("categoriePage")
	else
		session("categoriePage") = 1
		numPage = session("categoriePage")
	end if
end if
%>