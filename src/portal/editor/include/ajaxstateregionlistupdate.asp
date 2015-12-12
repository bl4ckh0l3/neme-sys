<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
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

	Dim field_name, field_val, objtype, id_objref
	field_val = request("field_val")

	Set objRef = New CountryClass
	if(field_val<>"")then
		On Error Resume Next
		Set specialFieldValue = objRef.findStateRegionListByCountry(field_val,"2,3")			    
		for each x in specialFieldValue
			key =  specialFieldValue(x).getStateRegionCode()%>
			<option value="<%=key%>"><%=Server.HTMLEncode(langEditor.getTranslated("portal.commons.select.option.country."&key))%></option>     
		<%next
		Set specialFieldValue = nothing
		if(Err.number <>0)then

		end if
	end if
	Set objRef = nothing
end if
%>
