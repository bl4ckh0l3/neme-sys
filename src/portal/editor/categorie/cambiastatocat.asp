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
	
	'/**
	'* Recupero tutti i parametri dal form e li elaboro
	'*/	
	Dim id_cat, stato_cat, items, page
	
	id_cat = request("id_cat_to_change")
	stato_cat = request("stato_cat")
	items = request("items")
	page = request("page")

	Dim objCat
	Set objCat = New CategoryClass	
	call objCat.changeCategoryVisibility(id_cat, stato_cat)
	Set objCat = nothing
	Set objUserLogged = nothing
	response.Redirect(Application("baseroot")&"/editor/categorie/Listacategorie.asp?items="&items&"&page="&page)				


	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>