<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
Dim userMail, pageURL, tellafriendMessage

'*** verifico se � stata passata la lingua dell'utente e la imposto come lang.setLangCode(xxx)
if not(isNull(request("lang_code"))) AND not(request("lang_code")="") AND not(request("lang_code")="null")  then
	lang.setLangCode(request("lang_code"))
	lang.setLangElements(lang.getListaElementsByLang(lang.getLangCode()))
end if

userMail = request("userMail")
pageURL = request("pageURL")
tellafriendMessage = request("tellafriendMessage")

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="<%="http://" & request.ServerVariables("SERVER_NAME") & Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
</head>
<body>
<div id="mailtellafriend">
<h2><%=lang.getTranslated("frontend.tellafriend.mail.label.intro")%></h2>
<br><br>
<%=lang.getTranslated("frontend.tellafriend.mail.label.user_email")%>:&nbsp;<%=userMail%><br><br>
<%=lang.getTranslated("frontend.tellafriend.mail.label.page_url")%>:&nbsp;<a href="<%=pageURL%>"><%=lang.getTranslated("frontend.tellafriend.mail.label.open_page")%></a><br><br>
<%=lang.getTranslated("frontend.tellafriend.mail.label.msg")%>:&nbsp;<%=tellafriendMessage%><br><br>
</div>
</body>
</html>
