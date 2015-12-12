<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<%
if not(isEmpty(Session("objCMSUtenteLogged"))) then
	On error resume next
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if
	
	'/**
	'* Recupero tutti i parametri dal form e li elaboro
	'*/	
	Dim id, lang_to_active, items, page
	
	id = request("id_lang_to_activate")
	lang_to_active = request("lang_to_active")
	items = request("items")
	page = request("page")
					
	Dim objLang, objSelLang
	Set objLang = New LanguageClass
	Set objSelLang = objLang.findLanguage(id)
	
	call objLang.updateActiveLanguage(id, lang_to_active)
	
	Set objSelLang = nothing
	Set objLang = nothing
	Set objUserLogged = nothing
	response.Redirect(Application("baseroot")&"/editor/language/InserisciLingua.asp?items="&items&"&page="&page)				


	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>