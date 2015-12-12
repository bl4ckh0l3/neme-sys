<%On Error Resume Next%>
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

Dim id_utente, active, action, resp, mailFriend
resp = 0
idFriend = request("id_utente")
mailFriend = objUsrLoggedTmp.findUserByID(idFriend).getEmail()
active = 0
action = 0
if not(request("active")="") then
	active = request("active")
end if

if not(request("action")="") then
	action = request("action")
end if


	if(action=0) then
		if (Cint(idFriend) <> -1 AND (Cint(idFriend) <> Cint(id_utente))) then
			if (objUsrLoggedTmp.bolHasFriend(idFriend, id_utente)) then
				resp="1"
			else
				resp="0"
			end if
		else
			resp="1"				
		end if
	elseif(action=1) then
		if (Cint(idFriend) <> -1 AND (Cint(idFriend) <> Cint(id_utente))) then
			if  not(objUsrLoggedTmp.bolHasFriend(idFriend, id_utente)) then
				Set objDB = New DBManagerClass
				Set objConn = objDB.openConnection()
				objConn.BeginTrans
				call objUsrLoggedTmp.insertFriendXUser(idFriend, id_utente, 1, objConn)
				call objUsrLoggedTmp.insertFriendXUser(id_utente, idFriend, active, objConn)

				if objConn.Errors.Count = 0 then
					objConn.CommitTrans	
					resp="1"
					
					Set objMail = New SendMailClass
					call objMail.sendMailCheckFriend(idFriend, mailFriend, id_utente, active, lang.getLangCode(), 1)
					Set objMail = Nothing
				else
					objConn.RollBackTrans
					resp="0"
				end if			
				Set objDB = nothing
			end if
		end if
	elseif(action=2) then
		if (Cint(idFriend) <> -1 AND (Cint(idFriend) <> Cint(id_utente))) then
			if (objUsrLoggedTmp.bolHasFriendActive(idFriend, id_utente) AND objUsrLoggedTmp.bolHasFriendActive(id_utente, idFriend)) then
				resp="1"
			else
				resp="0"
			end if
		else
			resp="0"				
		end if		
	end if

Set objUsrLoggedTmp = nothing

if(Err.number <> 0) then
	resp="0"
end if

response.write(resp)
%>