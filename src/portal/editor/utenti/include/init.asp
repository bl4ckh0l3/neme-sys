<%
if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp, msgMailSend
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
Set objUserLoggedTmp = nothing
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing


Public Sub sendMailCommunication(strMailBcc, strSubject, strHTML)
	on error resume next
	
	Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
	strServerName = "http://" & request.ServerVariables("SERVER_NAME")

	nomeDominio 				= Request.ServerVariables("HTTP_HOST")
	indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
	modulo						= Request.ServerVariables("HTTP_REFERER")
	browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

	'* creo gli oggetti cdosys sul server e li gestisco		
	DIM iMsg, Flds, iConf
	
	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
	
	if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
	else
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
	end if
	
	iConf.Fields.Update
	
	With iMsg
		Set .Configuration = iConf
			.To = Application("mail_sender")
			.From = Application("mail_sender")
			'.Sender = Application("mail_sender")
			'.cc = strMailBcc
			.bcc = strMailBcc
			.Subject = strSubject
			.HTMLBody =strHTML
			if not(Application("mail_server") = "") then
				.Send				
			end if
	End With

	if Err.number <> 0 then
		msgMailSend=langEditor.getTranslated("backend.utenti.lista.button.inserisci.mail_ko")
	end if	
End Sub
	
Dim objUtente, objListaUtenti, objListaRuoli, objUsrField, objListaField
Set objUtente = New UserClass
Set objUsrField = new UserFieldClass


Dim totPages, itemsXpageList, numPageList,itemsXpageField, numPageField

showTab="usrlist"
if(request("showtab")<>"")then
	showTab=request("showtab")
end if


if not(request("itemsList") = "") then
	session("listItems") = request("itemsList")
	itemsXpageList = session("listItems")
	session("listPage") = 1
else
	if not(session("listItems") = "") then
		itemsXpageList = session("listItems")
	else
		session("listItems") = 20
		itemsXpageList = session("listItems")
	end if
end if

if (showTab="usrlist") AND not(request("page") = "") then
	session("listPage") = request("page")
	numPageList = session("listPage")
else
	if not(session("listPage") = "") then
		numPageList = session("listPage")
	else
		session("listPage") = 1
		numPageList = session("listPage")
	end if
end if



if not(request("itemsField") = "") then
	session("fieldItems") = request("itemsField")
	itemsXpageField = session("fieldItems")
	session("fieldPage") = 1
else
	if not(session("fieldItems") = "") then
		itemsXpageField = session("fieldItems")
	else
		session("fieldItems") = 20
		itemsXpageField = session("fieldItems")
	end if
end if

if (showTab="usrfield") AND not(request("page") = "") then
	session("fieldPage") = request("page")
	numPageField = session("fieldPage")
else
	if not(session("fieldPage") = "") then
		numPageField = session("fieldPage")
	else
		session("fieldPage") = 1
		numPageField = session("fieldPage")
	end if
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

if(request("do_send_mail")="1")then
call sendMailCommunication(request("bcc_list"), request("mail_subject"), request("mail_body"))
if(msgMailSend="")then
msgMailSend=langEditor.getTranslated("backend.utenti.lista.button.inserisci.mail_ok")
end if
end if
%>