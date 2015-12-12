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
'* recupero i valori della news selezionata se id_lang <> -1
'*/

Dim objLanguage, objListaLanguage
Set objLanguage = New LanguageClass

Dim totPages, newsXpage, numPage

if not(request("items") = "") then
	session("languageItems") = request("items")
	itemsXpage = session("languageItems")
	session("languagePage") = 1
else
	if not(session("languageItems") = "") then
		itemsXpage = session("languageItems")
	else
		session("languageItems") = 20
		itemsXpage = session("languageItems")
	end if
end if

if not(request("page") = "") then
	session("languagePage") = request("page")
	numPage = session("languagePage")
else
	if not(session("languagePage") = "") then
		numPage = session("languagePage")
	else
		session("languagePage") = 1
		numPage = session("languagePage")
	end if
end if
%>