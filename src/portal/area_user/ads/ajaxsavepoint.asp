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
	
	Dim latitude,longitude,txtinfo,id_element, new_id, operation
	id = request("id")
	id_element = request("id_element")
	type_elem = request("type")
	latitude = request("latitude")
	longitude = request("longitude")
	txtinfo = request("txtinfo")
	operation = request("operation")
	
	Dim objfield
	Set objfield = new LocalizationClass
	if(operation="del")then
		call objfield.deletePoint(id)
	else
		call objfield.deletePoint(id)
		new_id = objfield.insertPointNoTransaction(id_element, type_elem, latitude, longitude, txtinfo)
		response.write(new_id)		
	end if			
	Set objfield = nothing
	Set objUserLogged = nothing
else
end if
%>