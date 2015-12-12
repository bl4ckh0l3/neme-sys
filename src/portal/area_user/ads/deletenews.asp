<%@ Language=VBScript %>
<% 
option explicit
On error resume next
%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->

<%
if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("guest_role"), 1) = 0) then
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
	
	call objLogger.write("cancellato contenuto --> id: "&id_news, objUserLogged.getUserName(), "info")
	
	Set objUserLogged = nothing	
	
	Set objLogger = nothing
	
	response.Redirect(Application("baseroot")&"/area_user/ads/ListaNews.asp")				


	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>