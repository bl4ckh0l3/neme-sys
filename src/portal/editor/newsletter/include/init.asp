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
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing

Dim objNewsletter, objListaNewsletter
Set objNewsletter = New NewsletterClass

Dim itemsXpage, numPage

if not(request("items") = "") then
	session("newsletterItems") = request("items")
	itemsXpage = session("newsletterItems")
	session("newsletterPage") = 1
else
	if not(session("newsletterItems") = "") then
		itemsXpage = session("newsletterItems")
	else
		session("newsletterItems") = 20
		itemsXpage = session("newsletterItems")
	end if
end if

if not(request("page") = "") then
	session("newsletterPage") = request("page")
	numPage = session("newsletterPage")
else
	if not(session("newsletterPage") = "") then
		numPage = session("newsletterPage")
	else
		session("newsletterPage") = 1
		numPage = session("newsletterPage")
	end if
end if	
%>