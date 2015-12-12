<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
Dim userMail, mailText, nome, cognome, telefono, indirizzo, zipcode, citta, nazione

'*** verifico se � stata passata la lingua dell'utente e la imposto come lang.setLangCode(xxx)
if not(isNull(request("lang_code"))) AND not(request("lang_code")="") AND not(request("lang_code")="null")  then
	lang.setLangCode(request("lang_code"))
	lang.setLangElements(lang.getListaElementsByLang(lang.getLangCode()))
end if

userMail= request("userMail")
mailText= request("mailText")
nome= request("nome")
cognome= request("cognome")
telefono= request("telefono")
indirizzo= request("indirizzo")
zipcode= request("zipcode")
citta= request("citta")
nazione= request("nazione")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=langEditor.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<link rel="stylesheet" href="<%="http://" & request.ServerVariables("SERVER_NAME") & Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
</head>
<body>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td class="tdContainerTop"><!-- #include virtual="/public/layout/include/header.inc" --></td>
  </tr>
  <tr>
    <td class="tdContainerContent">
	<table class="tableContent" border="0" align="left" cellpadding="0" cellspacing="0">
	  <tr>
		<td class="tdContentMailSend">
		<span class="labelMailSend"><%=langEditor.getTranslated("backend.utenti.mail.label.intro")%></span><br><br>
		<%=langEditor.getTranslated("backend.utenti.mail.label.intro_detail")%><br><br>
		
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.name")%>:</span>&nbsp;<%=nome%><br><br>
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.surname")%>:</span>&nbsp;<%=cognome%><br><br>
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.email")%>:</span>&nbsp;<%=userMail%><br><br>
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.telephone")%>:</span>&nbsp;<%=telefono%><br><br>
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.address")%>:</span>&nbsp;<%=indirizzo%><br><br>
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.zipcode")%>:</span>&nbsp;<%=zipcode%><br><br>
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.city")%>:</span>&nbsp;<%=citta%><br><br>
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.country")%>:</span>&nbsp;<%=nazione%><br><br>
		<span class="labelMailSend"><%=lang.getTranslated("backend.utenti.mail.label.message")%>:</span>&nbsp;<%=mailText%><br><br>

	</td>
	  </tr>
	</table>
	</td>
  </tr>
  <tr>
    <td class="tdContainerBott"><!-- include virtual="/public/layout/include/bottom.inc" --></td>
  </tr>
</table>
</body>
</html>
<%
Set objUtente = nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>
