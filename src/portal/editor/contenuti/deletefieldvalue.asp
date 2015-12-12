<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->

<%
if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if
	
	Dim id_field, value
	id_field = request("id_field")
	value = request("value")
	
	Dim objfield
	Set objfield = new ContentFieldClass
	call objfield.deleteContentFieldValueNoTransaction(id_field,value)				
	Set objfield = nothing
	Set objUserLogged = nothing
else
	'response.Redirect(Application("baseroot")&"/login.asp")
end if
%>