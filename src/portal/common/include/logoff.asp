<%On Error Resume Next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>LOG OFF</title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<%
'<!--nsys-demologoff1-->
'**** elimino l'utente loggato alla lista degli utenti online
for each x in onlineUsersList
	if onlineUsersList(x)=Session("objUtenteOnline") then
		onlineUsersList.remove(x)
	end if
next
if(onlineUsersList.Exists(Session.SessionID)= true) then
	onlineUsersList.remove(Session.SessionID)
end if
'<!---nsys-demologoff1-->

Session.Abandon()

Dim isHTTPS,strURL
strURL = Application("baseroot")&"/default.asp"
isHTTPS = Request.ServerVariables("HTTPS")
If isHTTPS = "on" Then
		strURL = "http://"&Request.ServerVariables("SERVER_NAME")&Application("baseroot")&"/default.asp"
end if
response.redirect(strURL)
%>
</body>
</html>