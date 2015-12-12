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

Dim objCurrency, objListaCurrency
Set objCurrency = New CurrencyClass

Dim totPages, newsXpage, numPage

if not(request("items") = "") then
	session("currencyItems") = request("items")
	itemsXpage = session("currencyItems")
	session("currencyPage") = 1
else
	if not(session("currencyItems") = "") then
		itemsXpage = session("currencyItems")
	else
		session("currencyItems") = 20
		itemsXpage = session("currencyItems")
	end if
end if

if not(request("page") = "") then
	session("currencyPage") = request("page")
	numPage = session("currencyPage")
else
	if not(session("currencyPage") = "") then
		numPage = session("currencyPage")
	else
		session("currencyPage") = 1
		numPage = session("currencyPage")
	end if
end if
%>