<%
if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
Set objUserLoggedTmp = nothing
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing

'/**
'* recupero i valori della news selezionata se id_target <> -1
'*/
Dim id_target, strDescrizione, iType, bolAutomatic
id_target = request("id_target")
strDescrizione = ""
iType = ""
bolAutomatic = 0

if (Cint(id_target) <> -1) then
	Dim objTarget, objSelTarget
	Set objTarget = New TargetClass
	Set objSelTarget = objTarget.findTarget(id_target)
	Set objTarget = nothing
	
	id_target = objSelTarget.getTargetID()
	strDescrizione = objSelTarget.getTargetDescrizione()		
	iType = objSelTarget.getTargetType()
	bolAutomatic = objSelTarget.isAutomatic()	
	Set objSelTarget = Nothing
end if
%>