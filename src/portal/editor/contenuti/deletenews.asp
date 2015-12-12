<%@ Language=VBScript %>
<% 
option explicit
On error resume next
%>
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
	
	Dim objLogger
	Set objLogger = New LogClass
	
	'/**
	'* Recupero tutti i parametri dal form e li elaboro
	'*/	
	Dim id_news	
	
	id_news = request("id_news_to_delete")
					
	Dim objNews
	Set objNews = New NewsClass
	call objNews.deleteNews(id_news)
	Set objNews = nothing

	'rimuovo l'oggetto dalla cache
	Set objCacheClass = new CacheClass
	call objCacheClass.remove("content-"&id_news)
	call objCacheClass.remove("listcf-"&id_news)
	call objCacheClass.removeByPrefix("findc", id_news)
	Set objCacheClass = nothing
			
	call objLogger.write("cancellato contenuto --> id: "&id_news, objUserLogged.getUserName(), "info")
	
	Set objUserLogged = nothing	
	
	Set objLogger = nothing
	
	response.Redirect(Application("baseroot")&"/editor/contenuti/ListaNews.asp")				


	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>