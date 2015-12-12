<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/SendMailClass.asp" -->
<%
Dim objUsrLogged, objUsrLoggedTmp, idUser

Set objUserLogged = new UserClass
if not(isEmpty(Session("objUtenteLogged"))) AND not(Session("objUtenteLogged") = "") then
	Set objUserLoggedTmp = objUserLogged.findUserByID(Session("objUtenteLogged"))
	if not(isNull(objUserLoggedTmp)) AND (Instr(1, typename(objUserLoggedTmp), "UserClass", 1) > 0) then  
		numIdUser = objUserLoggedTmp.getUserID()
	end if
	Set objUserLoggedTmp = nothing
	isLogged = true
elseif not(isEmpty(Session("objCMSUtenteLogged"))) AND not(Session("objCMSUtenteLogged") = "") then
	Set objCMSUtenteLoggedTmp = objUserLogged.findUserByID(Session("objCMSUtenteLogged"))
	if not(isNull(objCMSUtenteLoggedTmp)) AND (Instr(1, typename(objCMSUtenteLoggedTmp), "UserClass", 1) > 0) then
		numIdUser = objCMSUtenteLoggedTmp.getUserID() 
	end if
	Set objCMSUtenteLoggedTmp = nothing
	isLogged = true
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
Set objUserLogged = nothing

			
Dim id_element, element_type, objCommento, objSelectedCommento, del_commento, posted, id_commento, message, comment_type, active
id_element = request("id_element")
element_type = request("element_type")
id_commento = request("id_commento")
message = request("comment_message")
comment_type = request("comment_type")
active = request("active")
posted = 0

Set objCommento = New CommentsClass
			
if(message <> "") then
	if (isLogged) then
		Dim newCommentMax
		newCommentMax =  objCommento.insertCommentoNoTransaction(id_element, element_type, numIdUser, message, comment_type, active)
		posted = 1
		
		if(Application("use_comments_filter")=1 AND Application("mail_comment_receiver") <> "") then
			'Spedisco la mail di conferma registrazione
			Dim objMail
			Set objMail = New SendMailClass			
			call objMail.sendMailComment(newCommentMax, Application("mail_comment_receiver"), Application("str_editor_lang_code_default"))
			Set objMail = Nothing
		end if

	end if	
end if

Set objCommento = nothing

response.Redirect(Application("baseroot")&"/common/include/Controller.asp?posted="&posted&"&"&Request.QueryString())

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>