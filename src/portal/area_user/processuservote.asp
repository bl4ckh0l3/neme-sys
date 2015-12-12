<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUsrLogged, objUsrLoggedTmp, idFriend

Set objUserLogged = new UserClass
if not(isEmpty(Session("objUtenteLogged"))) AND not(Session("objUtenteLogged") = "") then
	Set objUserLoggedTmp = objUserLogged.findUserByID(Session("objUtenteLogged"))
	if not(isNull(objUserLoggedTmp)) AND (Instr(1, typename(objUserLoggedTmp), "UserClass", 1) > 0) then  
		idFriend = objUserLoggedTmp.getUserID()
	end if
	Set objUserLoggedTmp = nothing
	isLogged = true
elseif not(isEmpty(Session("objCMSUtenteLogged"))) AND not(Session("objCMSUtenteLogged") = "") then
	Set objCMSUtenteLoggedTmp = objUserLogged.findUserByID(Session("objCMSUtenteLogged"))
	if not(isNull(objCMSUtenteLoggedTmp)) AND (Instr(1, typename(objCMSUtenteLoggedTmp), "UserClass", 1) > 0) then
		idFriend = objCMSUtenteLoggedTmp.getUserID() 
	end if
	Set objCMSUtenteLoggedTmp = nothing
	isLogged = true
end if
Set objUserLogged = nothing

Dim id_utente,vote,message, id_usr_comment, comment_type
id_utente = request("id_utente")
vote = request("vote")
message = request("vote_message")
id_usr_comment = request("id_usr_comment")
comment_type = request("comment_type")

if (Cint(id_utente) <> -1 AND (Cint(idFriend) <> Cint(id_utente))) then
	Set objUserPreference = new UserPreferenceClass
	
	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()
	objConn.BeginTrans
	call objUserPreference.insertUserPreference(id_utente, idFriend, id_usr_comment, comment_type, vote, message, objConn)
						
	if objConn.Errors.Count = 0 then
		objConn.CommitTrans
	else
		objConn.RollBackTrans
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if			
	Set objDB = nothing	
	
	Set objUserPreference = nothing

	response.Redirect(Application("baseroot")&"/common/include/Controller.asp?vode_done=1&"&Request.QueryString())	
else
	response.Redirect(Application("baseroot")&"/common/include/Controller.asp?vode_done=0&"&Request.QueryString())				
end if

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>