<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->

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
	
	Dim id_target, strDescrizione, iType, bolDeltArget, bolAutomatic
	id_target = request("id_target")
	strDescrizione = request("descrizione")
	bolAutomatic = request("automatic")
	iType = request("target_type")
	bolDeltArget = request("delete_target")
	
	Dim objTarget
	Set objTarget = New TargetClass
	
	if (Cint(id_target) <> -1) then
		if(strComp(bolDeltArget, "del", 1) = 0) then
			if(objTarget.findTargetAssociations(id_target)) then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=011")		
			else
				call objTarget.deleteTarget(id_target)
				response.Redirect(Application("baseroot")&"/editor/targets/ListaTarget.asp")	
			end if
		
		end if
		
	
		call objTarget.modifyTargetNoTransaction(id_target, strDescrizione, iType, 0, bolAutomatic)
		Set objTarget = nothing
		response.Redirect(Application("baseroot")&"/editor/targets/ListaTarget.asp")		
	else
		call objTarget.insertTargetNoTransaction(strDescrizione, iType, 0, bolAutomatic)
		Set objTarget = nothing
		response.Redirect(Application("baseroot")&"/editor/targets/ListaTarget.asp")				
	end if

	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>