<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUsrLogged, objUsrLoggedTmp, idFriend
Set objUsrLoggedTmp = new UserClass
Set objUsrLogged = objUsrLoggedTmp.findUserByID(Session("objUtenteLogged"))
id_utente = objUsrLogged.getUserID()
Set objUsrLogged = nothing%>
<%

Dim id_utente,vote
idFriend = request("id_utente")

if (Cint(idFriend) <> -1 AND (Cint(idFriend) <> Cint(id_utente))) then		
	if  not(objUsrLoggedTmp.bolHasFriend(idFriend, id_utente)) then
		response.Redirect(Application("baseroot") & "/area_user/friendlist.asp?add_done=0&id_utente="&id_utente)
	end if

	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()
	objConn.BeginTrans
	call objUsrLoggedTmp.deleteFriendXUser(id_utente, idFriend, objConn)
	call objUsrLoggedTmp.deleteFriendXUser(idFriend, id_utente, objConn)

	if objConn.Errors.Count = 0 then
		objConn.CommitTrans
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