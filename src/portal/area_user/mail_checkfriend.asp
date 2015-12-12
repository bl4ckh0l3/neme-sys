<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%

Dim idFriend, action

'*** verifico se è stata passata la lingua dell'utente e la imposto come lang.setLangCode(xxx)
if not(isNull(request("lang_code"))) AND not(request("lang_code")="") AND not(request("lang_code")="null")  then
	lang.setLangCode(request("lang_code"))
	lang.setLangElements(lang.getListaElementsByLang(lang.getLangCode()))
end if

idFriend = request("idUtente")
action = request("action")
active = request("active")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
</head>
<body>
	<div>
	<h2><%=lang.getTranslated("frontend.confirm_friend.mail.label.intro_checkfriend")%></h2>
	<br><br>


	<%
	Dim  objUser, objFriend
	Set objUser = New UserClass
	Set objFriend = objUser.findUserByID(idFriend)%>
	<div style="float:left;padding-right:10px;">
	<%if(objUser.hasImageUser(idFriend))then%>
	<img class="imgAvatarUserOn" src="<%="http://" & request.ServerVariables("SERVER_NAME") &Application("baseroot") & "/common/include/userImage.asp?userID="&idFriend%>" <%If (InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE",1) > 0) then response.write(" width=""50"" height=""50""") end if%> />
	<%else%>
	<img class="imgAvatarUserOn" src="<%="http://" & request.ServerVariables("SERVER_NAME") &Application("baseroot") & "/common/img/unkow-user.jpg"%>" <%If (InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE",1) > 0) then response.write(" width=""50"" height=""50""") end if%> />
	<%end if%>
	</div>
	<div style="padding-bottom:15px;"><%="<i>"&objFriend.getUserName()&"</i>"%></div>
	<p align="center">	
	<%if(action=0) then
		if(active=1)then%>		
			<%=lang.getTranslated("frontend.confirm_friend.mail.label.friend_confirmed")%>		
		<%else%>
			<%=lang.getTranslated("frontend.confirm_friend.mail.label.friend_noconfirmed")%>		
		<%end if
	else%>
		<%=lang.getTranslated("frontend.confirm_friend.mail.label.friend_askadd")%>
	<%end if	
	Set objFriend = nothing
	Set objUser = nothing%>
	</p>
	<hr><br/><br/><a href="<%="http://" & request.ServerVariables("SERVER_NAME") &Application("baseroot") &"/area_user/manageUser.asp"%>"><%=lang.getTranslated("backend.confirm_comment.mail.label.confirm_friend")%></a><br/><br/>

	</div>
</body>
</html>