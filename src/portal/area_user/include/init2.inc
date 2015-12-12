<%
Dim id_utente, strUserName, strPwd, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, objUsrTarget
Dim bolUserActive, numSconto, strAdminComments
Dim dateInsertDate, dateModifyDate, objNewsletter, objNewsletterUsr
Dim toDo, usrHasImg
Dim bolPublic, numUserGroup
Dim default_date
default_date = "dd/mm/yyyy"
toDo = request("toDo")

id_utente = -1
strUserName = lang.getTranslated("frontend.area_user.manage.label.username") 
strPwd = ""
strEmail = lang.getTranslated("frontend.area_user.manage.label.email")
strUsrRuolo = Application("guest_role")
bolPrivacy = true
bolNewsletter = true 
if(Application("confirm_registration") = 1) then
bolUserActive = "1"
else
bolUserActive = "0"
end if
numSconto = 0
strAdminComments = ""
objUsrTarget = null
dateInsertDate = ""
dateModifyDate = ""
objNewsletterUsr = null
usrHasImg = false
bolPublic = "0"
numUserGroup = null
'<!--nsys-ausrinc1-->
Dim objGroup
Set objGroup = New UserGroupClass
if(strComp(typename(objGroup.findUserGroupDefault()), "UserGroupClass", 1) = 0)then
	numUserGroup = objGroup.findUserGroupDefault().getID()
end if
Set objGroup = nothing
'<!---nsys-ausrinc1-->

Set objNewsletter = new NewsletterClass


if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp, objListaRuoli
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objListaRuoli = objUserLoggedTmp.getListaRuoli()
	Set objUserLoggedTmp = nothing
		
	id_utente = objUserLogged.getUserID()
	strUserName = objUserLogged.getUserName()
	strPwd = objUserLogged.getPassword()
	strEmail = objUserLogged.getEmail()
	strUsrRuolo = objUserLogged.getRuolo()
	bolPrivacy = objUserLogged.getPrivacy()
	bolNewsletter = objUserLogged.getNewsletter() 
	bolUserActive = objUserLogged.getUserActive()
	numSconto = objUserLogged.getSconto()
	strAdminComments = objUserLogged.getAdminComments()
	dateInsertDate = objUserLogged.getInsertDate()
	dateModifyDate = objUserLogged.getModifyDate()
	usrHasImg = objUserLogged.hasImageUser(id_utente)
	bolPublic = objUserLogged.getPublic()
  numUserGroup = objUserLogged.getGroup()

	'/*
	' * se l'utente loggato è admin lo mando all'inserimento utenti del content manager
	' */
	if(strComp(Cint(strUsrRuolo), Application("admin_role"), 1) = 0) then
		response.Redirect(Application("baseroot") & "/editor/InserisciUtente.asp?id_utente="& id_utente)
	end if
	
	if not(isNull(objUserLogged.getTargetPerUser(id_utente))) then
		Set objUsrTarget = objUserLogged.getTargetPerUser(id_utente)
	end if		

	if not(isNull(objUserLogged.getNewsletterPerUser(id_utente))) then
		Set objNewsletterUsr = objUserLogged.getNewsletterPerUser(id_utente)
	end if
elseif not(toDo = "") AND (strComp(toDo, "modify", 1)= 0) then
	response.Redirect(Application("baseroot")&"/login.asp?from=modify")
end if



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

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>