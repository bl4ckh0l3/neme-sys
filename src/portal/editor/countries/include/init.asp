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

Dim objCountry, objListaCountry
Set objCountry = New CountryClass

Dim totPages, itemsXpage, numPage

if not(request("items") = "") then
	session("countryItems") = request("items")
	itemsXpage = session("countryItems")
	session("countryPage") = 1
else
	if not(session("countryItems") = "") then
		itemsXpage = session("countryItems")
	else
		session("countryItems") = 20
		itemsXpage = session("countryItems")
	end if
end if

if not(request("page") = "") then
	session("countryPage") = request("page")
	numPage = session("countryPage")
else
	if not(session("countryPage") = "") then
		numPage = session("countryPage")
	else
		session("countryPage") = 1
		numPage = session("countryPage")
	end if
end if

Dim search_key
	
if not(Trim(request("search_key")) = "") then
	'** sostituisco: טיאעשל'
	'** con: &egrave;&eacute;&agrave;&ograve;&ugrave;&igrave;&#39;
	tmp_key = Trim(request("search_key"))
	'tmp_key = Replace(tmp_key, "ט", "&egrave;", 1, -1, 1)
	'tmp_key = Replace(tmp_key, "י", "&eacute;", 1, -1, 1)
	'tmp_key = Replace(tmp_key, "א", "&agrave;", 1, -1, 1)
	'tmp_key = Replace(tmp_key, "ע", "&ograve;", 1, -1, 1)
	'tmp_key = Replace(tmp_key, "ש", "&ugrave;", 1, -1, 1)
	'tmp_key = Replace(tmp_key, "ל", "&igrave;", 1, -1, 1)
	'tmp_key = Replace(tmp_key, "'", "&#39;", 1, -1, 1)
						
	session("search_key") = tmp_key
	search_key = session("search_key")
else
	if not(session("search_key") = "") then
		search_key = session("search_key")
	else
		session("search_key") = ""
		search_key = session("search_key")
	end if
end if

if(not(isNull(request("resetMenu"))) AND request("resetMenu") = "1") then
	session("countryPage") = 1
	numPage = session("countryPage")
	session("search_key") = ""
	search_key = session("search_key")
end if
%>