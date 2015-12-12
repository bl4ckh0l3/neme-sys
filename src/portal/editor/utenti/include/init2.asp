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
Dim id_utente, strUserName, strPwd, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, objUsrTarget
Dim bolUserActive, numSconto, strAdminComments, bolPublic, numUserGroup
Dim dateInsertDate, dateModifyDate, objNewsletter, objNewsletterUsr

Dim objListaTarget, bolExistTarget, objTargets
Set objTargets = new TargetClass
Set objListaTarget = objTargets.getListaTarget()
Set objTargets = nothing

id_utente = request("id_utente")
strUserName = "" 
strPwd = ""
strEmail = ""
strUsrRuolo = 3
bolPrivacy = true
bolNewsletter = false 
objUsrTarget = null
if(Application("confirm_registration") = 1) then
bolUserActive = "1"
else
bolUserActive = "0"
end if
numSconto = 0
strAdminComments = ""
dateInsertDate = ""
dateModifyDate = ""
objNewsletterUsr = null
bolPublic = "0"
numUserGroup = null

'<!--nsys-usrinc1-->
Dim objGroup
Set objGroup = New UserGroupClass
if(strComp(typename(objGroup.findUserGroupDefault()), "UserGroupClass", 1) = 0)then
	numUserGroup = objGroup.findUserGroupDefault().getID()
end if
'<!---nsys-usrinc1-->

Dim objUtente, objSelUtente, objListaRuoli
Set objUtente = New UserClass
Set objListaRuoli = objUtente.getListaRuoli()
Set objNewsletter = new NewsletterClass

if (Cint(id_utente) <> -1) then
	Set objSelUtente = objUtente.findUserByID(id_utente)
	
	id_utente = objSelUtente.getUserID()
	strUserName = objSelUtente.getUserName()
	strPwd = objSelUtente.getPassword()
	strEmail = objSelUtente.getEmail()
	strUsrRuolo = objSelUtente.getRuolo()
	bolPrivacy = objSelUtente.getPrivacy()
	bolNewsletter = objSelUtente.getNewsletter() 
	numSconto = objSelUtente.getSconto()
	bolUserActive = objSelUtente.getUserActive()
	strAdminComments = objSelUtente.getAdminComments()
	bolPublic = objSelUtente.getPublic()
	numUserGroup = objSelUtente.getGroup()
	dateInsertDate = objSelUtente.getInsertDate()
	dateModifyDate = objSelUtente.getModifyDate()

	if not(isNull(objSelUtente.getTargetPerUser(id_utente))) then
		Set objUsrTarget = objSelUtente.getTargetPerUser(id_utente)
	end if			

	if not(isNull(objSelUtente.getNewsletterPerUser(id_utente))) then
		Set objNewsletterUsr = objSelUtente.getNewsletterPerUser(id_utente)
	end if
end if

Set objUtente = nothing

'<!--nsys-usrinc2-->
'********** RECUPERO LA LISTA DI GRUPPI UTENTE DISPONIBILI
Dim objDispGroup
On Error Resume Next
Set objDispGroup = objGroup.getListaUserGroup()
if(Err.number <> 0) then
end if
Set objGroup = nothing
'<!---nsys-usrinc2-->


'********** RECUPERO LA LISTA DI FIELD UTENTE DISPONIBILI
Dim objUserField, objListUserField, hasUserFields
hasUserFields=false
On Error Resume Next
Set objUserField = new UserFieldClass
Set objListUserField = objUserField.getListUserField(1,"1,3")
if(objListUserField.count > 0)then
	hasUserFields=true
end if
if(Err.number <> 0) then
	hasUserFields=false
end if
%>