<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->

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
	
	Dim action
	action = request("action")
	
	Dim id_field,id_group, description, id_type,id_type_content, values, order, required, enabled, maxLenght, useFor, bolDelField	
	id_field = request("id_field")
	id_group = request("id_group")
	description = request("description")
	id_type = request("id_type")
	id_type_content = request("id_type_content")
	values = request("field_values")
	order = request("order")
	required = request("required")
	enabled = request("enabled")
	maxLenght = request("max_lenght")
	useFor = request("use_for")
	bolDelField = request("delete_field")
	
	Dim desc_new_group, order_new_group, id_del_group
	desc_new_group = request("desc_new_group")
	order_new_group = request("order_new_group")
	id_del_group = request("id_del_group")
	
	Dim objfield, objGroup
	Set objfield = new UserFieldClass
	Set objGroup = New UserFieldGroupClass

	if (strComp(action, "del_group", 1) = 0) then
		call objGroup.deleteUserFieldGroup(id_del_group)
		Set objGroup = nothing
		response.Redirect(Application("baseroot")&"/editor/utenti/inserisciField.asp?id_field="&id_field)	
	elseif(strComp(action, "ins_group", 1) = 0) then
		call objGroup.insertUserFieldGroupNoTransaction(desc_new_group,order_new_group)		
		Set objGroup = nothing
		response.Redirect(Application("baseroot")&"/editor/utenti/inserisciField.asp?id_field="&id_field)
	else
		if (Cint(id_field) <> -1) then
			if(strComp(bolDelField, "del", 1) = 0) then
				call objfield.deleteUserField(id_field)
				response.Redirect(Application("baseroot")&"/editor/utenti/ListaUtenti.asp?showtab=usrfield")	
			end if
			
			call objfield.modifyUserFieldNoTransaction(id_field,description, id_group, order,id_type,id_type_content,values,required,enabled,Trim(maxLenght),useFor)
			Set objfield = nothing
			response.Redirect(Application("baseroot")&"/editor/utenti/ListaUtenti.asp?showtab=usrfield")		
		else
			call objfield.insertUserFieldNoTransaction(description, id_group, order,id_type,id_type_content,values,required,enabled,Trim(maxLenght),useFor)
			Set objfield = nothing
			response.Redirect(Application("baseroot")&"/editor/utenti/ListaUtenti.asp?showtab=usrfield")				
		end if
	end if

	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>