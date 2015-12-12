﻿<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/SendMailClass.asp" -->
<%if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUsrLogged, objUsrLoggedTmp, idFriend
Set objUsrLoggedTmp = new UserClass
Set objUsrLogged = objUsrLoggedTmp.findUserByID(Session("objUtenteLogged"))
id_utente = objUsrLogged.getUserID()
Set objUsrLogged = nothing%>
<%

Dim id_utente,active, mailFriend
idFriend = request("id_utente")
mailFriend = objUsrLoggedTmp.findUserByID(idFriend).getEmail()
active = 0
if not(request("active")="") then
	active = request("active")
end if

if (Cint(idFriend) <> -1 AND (Cint(idFriend) <> Cint(id_utente))) then	
	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()
	objConn.BeginTrans
	call objUsrLoggedTmp.updateFriendStatus(idFriend, id_utente, active, objConn)

	if objConn.Errors.Count = 0 then
		objConn.CommitTrans
		
		Set objMail = New SendMailClass
		call objMail.sendMailCheckFriend(idFriend, mailFriend, id_utente, active, lang.getLangCode(), 0)
		Set objMail = Nothing
	else
		objConn.RollBackTrans
		response.Redirect(Application("baseroot") & "/area_user/friendlist.asp?add_done=0&id_utente="&id_utente)
	end if			
	Set objDB = nothing	

	response.Redirect(Application("baseroot") & "/area_user/friendlist.asp?add_done=1&id_utente="&id_utente)	
else
	response.Redirect(Application("baseroot") & "/area_user/friendlist.asp?add_done=0&id_utente="&id_utente)				
end if

Set objUsrLoggedTmp = nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>