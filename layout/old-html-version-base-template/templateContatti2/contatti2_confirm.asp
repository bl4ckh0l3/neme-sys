<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/SendMailClass.asp" -->

<%
Dim strGerarchia, strCSS

strGerarchia = request("gerarchia")

Dim sendmail, userMail, mailText, boolMailSent
Dim nome, cognome, telefono, indirizzo, cap, citta, nazione, infoBy
mailTo = request("mailTo")
userMail = request("email")
mailText =  request("testo")
nome = request("nome")
cognome = request("cognome")
telefono = request("telefono")
indirizzo = request("indirizzo")
cap = request("cap")
citta = request("citta")
nazione = request("nazione")
infoBy = request("infoBy")

Set objMail = New SendMailClass
call objMail.sendMailContactExtended(mailTo, userMail, mailText, nome, cognome, telefono, indirizzo, cap, citta, nazione, infoBy)
Set objMail = Nothing
%>
<!-- #include virtual="/common/include/setTemplateTargetList.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Testa Denis; email:blackhole01@gmail.com">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="<%=Application("baseroot") & "/common/css/stile.css"%>" type="text/css">
<%if not(isNull(strCSS)) ANd not(strCSS = "") then%>
<link rel="stylesheet" href="<%=Application("baseroot") & strCSS%>" type="text/css">
<%end if%>
<SCRIPT SRC="<%=Application("baseroot") & "/common/js/javascript_global.js"%>"></SCRIPT>
</head>
<body>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
    <td class="tdContainerTop"><!-- #include file="include/header.inc" --></td>	
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td class="tdContainerContentConfirm">
		<table border="0" cellpadding="0" cellspacing="0" align="left">
		  <tr>
			<td class="contatti2Left">&nbsp;</td>
			<td><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateContatti2/img/sfondo_center.jpg"%>" vspace="0" hspace="0" border="0"></td>
			<td class="contatti2RightConfirm"><%=lang.getTranslated("frontend.template_contatti.label.mail_sent")%></td>
		  </tr>
		</table>
	</td>
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td>
	<table border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="left" valign="top"><!--<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateContatti2/img/bottom_title.gif"%>" vspace="0" hspace="0" border="0">--></td>
	<td align="left" valign="top">
	<!--
	<img src="<%'=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="20" align="top"><br/>
	<span class="bottCopyright"><%'=lang.getTranslated("frontend.template_contatti.label.testo_bottom_copyright")%></span><br/>
	<img src="<%'=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="10" align="top"><br/>
	<span class="bottOrari"><%'=lang.getTranslated("frontend.template_contatti.label.testo_bottom_orari")%></span>
	-->
	</td>
	</tr>
	</table>
	</td>	
  </tr>
</table>
</body>
</html>
<%
if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>