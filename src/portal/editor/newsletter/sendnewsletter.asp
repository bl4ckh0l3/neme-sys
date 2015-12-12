<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<%
if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if
	
	Dim objNewsLetter, idnews, objNewsletterNews, objNews, choose_newsletter
	idnews = request.QueryString("id_news")
	choose_newsletter = request.QueryString("choose_newsletter")
	
	Set objNews = new NewsClass
	Set objNewsLetter = new NewsletterClass
	Set objNewsletterNews = objNews.findNewsByID(idnews)
	
	call objNewsLetter.sendNewsletter(objNewsletterNews, choose_newsletter)
	Set objNewsLetter = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>