<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/SendMailClass.asp" -->

<%
Dim strGerarchia, strCSS, News

strGerarchia = request("gerarchia")

Set News = New NewsClass

Dim sendmail, userMail, mailText, boolMailSent, objMail
Dim nome, cognome, telefono, indirizzo, cap, citta, nazione, infoBy, azienda, pivacf, web, listOfItemsInput
Dim objSelNews, listOfItem, subList, currID, currQta
Dim totQta

mailTo = request("mailTo")
userMail = request("email")
mailText =  request("testo")
nome = request("nome")
cognome = request("cognome")
azienda = request("azienda")
'pivacf = request("pivacf")
telefono = request("telefono")
indirizzo = request("indirizzo")
cap = request("cap")
citta = request("citta")
nazione = request("nazione")
infoBy = request("infoBy")
web = request("web")
listOfItemsInput = request("listOfItemsInput")
thisCatDescDataScad = request("thisCatDescDataScad")

Set objMail = New SendMailClass

listOfItem = Split(Left(Session("listenozzeCarrello"),Len(Session("listenozzeCarrello"))-1), "|", -1, 1)

for y=LBound(listOfItem) to UBound(listOfItem)
	subList = Split(listOfItem(y), "#", -1, 1)
	currID = subList(0)
	currQta = subList(1)
	Set objSelNews = News.findNewsByID(currID)	
	totQta = Cint(objSelNews.getAbstract1()) - Cint(currQta)
	
	if(totQta >= 0) then	
		call News.modifyNewsNoTransaction(objSelNews.getNewsID(), objSelNews.getTitolo(), totQta, objSelNews.getAbstract2(), objSelNews.getAbstract3(), objSelNews.getTesto(), objSelNews.getDataInsNews(), objSelNews.getDataPubNews(), objSelNews.getDataDelNews(), objSelNews.getStato())
	else
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=028")
	end if
	
	Set objSelNews = nothing
next

call objMail.sendMailContactListenozze(mailTo, userMail, mailText, nome, cognome, azienda, telefono, indirizzo, cap, citta, nazione, infoBy, web, listOfItemsInput, thisCatDescDataScad)

Session("listenozzeCarrello") = ""
Set objMail = Nothing
Set News = Nothing
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
<script language="JavaScript">
<!--
function refreshParent() {
	window.opener.location.href = window.opener.location.href;
	
	/*
	if (window.opener.progressWindow){
		window.opener.progressWindow.close()
	}
	window.close();
	*/
}
//-->
</script>
</head>
<body onLoad="refreshParent();">
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
			<td class="confirmPULeft">&nbsp;</td>
			<td><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/liste_nozze.jpg"%>" vspace="0" hspace="0" border="0"></td>
			<td class="confirmPURightConfirm"><%=lang.getTranslated("frontend.template_listenozze.label.mail_sent")%></td>
		  </tr>
		</table>
	</td>
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td>&nbsp;</td>	
  </tr>
</table>
</body>
</html>
<%
if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>