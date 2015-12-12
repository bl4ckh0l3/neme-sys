<!-- #include file="Objects/DBManagerClass.asp" -->
<!-- #include file="Objects/UserClass.asp" -->
<!-- #include file="Objects/GUIDClass.asp" -->
<!-- #include file="Objects/LanguageClass.asp" -->
<!-- #include file="Objects/SendMailClass.asp" -->
<%
'** imposto il protocollo http o https da usare in tutte le pagine all'occorrenza			
base_url = "http://"
Dim isInitHTTPS
isInitHTTPS = Request.ServerVariables("HTTPS")
If isInitHTTPS = "on" AND Application("use_https") = 1 Then
	base_url = "https://"
End If		


'*************************** CODICE DI PROTEZIONE DA SQL INJECTION DA USARE IN TUTTE LE PAGINE
dim regEx, item
set regEx = New RegExp
regEx.Pattern = "banner82|nihaorr1|adw95|xp_|;|--|/\*|<script|</script|ntext|nchar|nvarchar|alter|begin|create|cursor|declare|delete|drop|exec|execute|fetch|insert|kill|open|sys|sysobjects|syscolumns|table|update|varchar|select"
regEx.IgnoreCase = true
regEx.Multiline = true
For Each item in Request.QueryString
    If regEx.test(Request.QueryString(item)) Then
        Response.redirect(base_url&Request.ServerVariables("SERVER_NAME")&Application("baseroot")&Application("error_page")&"?error=032")
    End IF
Next

For Each item in Request.Form
    If regEx.test(Request.Form(item)) Then
        Response.redirect(base_url&Request.ServerVariables("SERVER_NAME")&Application("baseroot")&Application("error_page")&"?error=032")
    End IF
Next


Dim objUtente, objUtenteVerify, strParamUser, strParamPwd, strParamMail, paramFrom, sPreviousURL, keepLogged, userCookie
paramFrom = request("from")

if(paramFrom = "") then
	paramFrom = "default"
end if

'sPreviousURL = Request.ServerVariables("HTTP_REFERER")
'if(sPreviousURL = "") then
'	sPreviousURL = Application("baseroot")&"/default.asp"
'end if

keepLogged = request("keep_logged")
strParamUser = request("j_username")
strParamPwd = request("j_password")
strParamMail = request("j_email")
	
Set objUtente = new UserClass

if not(paramFrom = "") AND (strComp(paramFrom, "lost_pwd", 1) = 0) then
	Set objUtenteVerify = objUtente.findUserByUserAndMail(strParamUser, strParamMail)	

	'**** CREO IL GUID PER IL NUOVA PASSWORD
	Dim objGUID, strGUID
	Set objGUID = new GUIDClass
	strGUID = objGUID.CreatePasswordGUID()
	Set objGUID = nothing

	if not (isNull(objUtenteVerify)) then

		'Aggiorno password su DB
		call objUtente.changePassword(objUtenteVerify.getUserID(), strGUID)

		'recupero la lingua da utilizzare per la mial, se non presente uso il default
		langMail = Application("str_lang_code_default")
		if(request("lang_mail")<>"") then
			langMail = request("lang_mail")
		end if

		'Spedisco la mail di recupero password
		Dim objMail
		Set objMail = New SendMailClass
		call objMail.sendMailUserPwd(objUtenteVerify.getUserID(), strGUID, strParamMail, langMail)
		Set objMail = Nothing
	
		response.redirect(Application("baseroot")&"/login.asp?message=001")
	else
		response.redirect(Application("baseroot")&"/login.asp?error=002")		
	end if

	
else
	Set objUtenteVerify = objUtente.login(strParamUser, strParamPwd, paramFrom)
end if

if not (isNull(objUtenteVerify)) then		

	if(strComp(Cint(objUtenteVerify.getRuolo()), Application("guest_role"), 1) = 0) then
		Session("objUtenteLogged") = objUtenteVerify.getUserID()
'<!--nsys-demoverify1-->
		Session("objUtenteOnline") = objUtenteVerify.getUserID()&"|"&objUtenteVerify.getPublic()&"|"&objUtenteVerify.hasImageUser(objUtenteVerify.getUserID())&"|"&objUtenteVerify.getUserName()
		
		On Error Resume Next
		'**** aggiungo l'utent eloggato alla lista degli utenti online
		for each x in onlineUsersList			
			if onlineUsersList(x)=Session("objUtenteOnline") then
				onlineUsersList.remove(x)
			end if
		next
		if(onlineUsersList.Exists(Session.SessionID)) then
			onlineUsersList.remove(Session.SessionID)
		end if			
		onlineUsersList.add Session.SessionID, Session("objUtenteOnline")

		'response.cookies(Application("srt_default_server_name"))("user_online")=Session("objUtenteOnline")
		'response.cookies(Application("srt_default_server_name")).Expires=DateAdd("m",6,date())
			
		if(Err.number <>0) then			
		end if
'<!---nsys-demoverify1-->
		
		'**** GESTISCO LA SCRITTURA DEL COOKIE PER MANTENERE L'UTENTE LOGGATO
		if(keepLogged = "1") then
			response.cookies(Application("srt_default_server_name"))("id_user")=objUtenteVerify.getUserID()
			response.cookies(Application("srt_default_server_name")).Expires=DateAdd("m",6,date())		
		end if
		
		select Case paramFrom
		Case "modify"
			response.redirect(Application("baseroot")&"/area_user/manageUser.asp")
		Case "lost_pwd"
			response.redirect(Application("baseroot")&"/area_user/manageUser.asp")
		Case "area_user"
			'response.redirect(Application("baseroot")&"/area_user/manageuser.asp")
			response.redirect(Application("baseroot")&"/default.asp")
		Case "carrello"
			response.redirect(Application("baseroot")&Application("dir_upload_templ")&"shopping-card/card.asp")
		Case "default"
			response.redirect(Application("baseroot")&"/default.asp")
		Case Else
			response.redirect(paramFrom)	
		End Select	
	else
		Session("objCMSUtenteLogged") = objUtenteVerify.getUserID()
		
		'**** GESTISCO LA SCRITTURA DEL COOKIE PER MANTENERE L'UTENTE LOGGATO
		if(keepLogged = "1") then
			response.cookies(Application("srt_default_server_name"))("id_bo")=objUtenteVerify.getUserID()
			response.cookies(Application("srt_default_server_name")).Expires=DateAdd("m",6,date())			
		end if

		response.redirect(Application("baseroot")&"/editor/index.asp?cssClass=AH")	
	end if	
end if
Set objUtenteVerify = nothing
Set objUtente = nothing
%>