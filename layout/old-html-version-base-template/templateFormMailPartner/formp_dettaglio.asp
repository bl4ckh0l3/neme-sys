<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->

<%
Dim id_news, News, objCurrentNews, strGerarchia, strCSS, objFileXNews, objListaFilePerNews
Dim objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim totPages, newsXpage, numPage

Dim stato
stato = 1
if(request("is_preview_content")) then
	stato = null
end if

Set News = New NewsClass
id_news = request("id_news")
strGerarchia = request("gerarchia")
order_news_by = 2
newsXpage = 5
numPage = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if


Dim destination, sendmail, userMail, mailObject, mailText, boolMailSent
destination = request("destination")
sendmail = request("sendmail")
userMail = request("email")
mailObject = request("oggetto")
mailText =  request("testo")
boolMailSent = false

if(sendmail = 1) then
	Set objMail = New SendMailClass
	call objMail.sendMailContact(userMail, mailObject, mailText, destination)
	Set objMail = Nothing	
	boolMailSent = true
		
end if
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
function sendMail(){
	if(controllaCampiMail()){
		document.form_send_mail.submit();
	}else{
		return;
	}
}

function controllaCampiMail(){
	var strMail = document.form_send_mail.email.value;
	if(strMail != ""){
		if (strMail.indexOf("@")<2 || strMail.indexOf(".")==-1 || strMail.indexOf(" ")!=-1 || strMail.length<6){
			alert("<%=lang.getTranslated("frontend.template_form.js.alert.wrong_mail")%>");
			document.form_send_mail.email.focus();
			return false;
		}
	}else if(strMail == ""){
		alert("<%=lang.getTranslated("frontend.template_form.js.alert.insert_mail")%>");
		document.form_send_mail.email.focus();
		return false;
	}	
	
	if(document.form_send_mail.oggetto.value == ""){
		alert("<%=lang.getTranslated("frontend.template_form.js.alert.insert_oggetto")%>");
		document.form_send_mail.oggetto.focus();
		return false;
	}	
	
	if(document.form_send_mail.testo.value == ""){
		alert("<%=lang.getTranslated("frontend.template_form.js.alert.insert_testo")%>");
		document.form_send_mail.testo.focus();
		return false;
	}
	
	return true;
}
</script>
</head>
<body>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td class="tdContainerTop">
	<!-- #include virtual="/common/include/header.inc" -->	
	</td>
  </tr>
  <tr>
    <td class="tdContainerContent">
	<table class="tableContent" border="0" align="left" cellpadding="0" cellspacing="0">
	  <tr>
		<td class="tdMenu"><!-- #include virtual="/common/include/MenuFruizione.inc" --></td>
		<td class="tdContent"><!-- #include virtual="/common/include/MenuTips.inc" -->
		<table class="tableFormDettaglio" width="0" border="0" cellpadding="0" cellspacing="0">
		<tr><td valign="top">
		<%	
		Dim bolHasObj
		bolHasObj = false
		
		on error Resume Next
		if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
			Set objListaNews = News.findNews(null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_news_by)	
			
			if(objListaNews.Count > 0) then
				Dim objSelNews, newsCounter, iIndex, objTmpNews, FromNews, ToNews, Diff
				iIndex = objListaNews.Count
				FromNews = ((numPage * newsXpage) - newsXpage)
				Diff = (iIndex - ((numPage * newsXpage)-1))
				if(Diff < 1) then
					Diff = 1
				end if
				
				ToNews = iIndex - Diff
				
				totPages = iIndex\newsXpage
				if(totPages < 1) then
					totPages = 1
				elseif((iIndex MOD newsXpage <> 0) AND not ((totPages * newsXpage) >= iIndex)) then
					totPages = totPages +1	
				end if		
			
				bolHasObj = true
			end if
		end if
			
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			bolHasObj = false
		end if				
		
		if((isNull(id_news)) OR id_news = "" ) AND bolHasObj then
			Dim objTmpCurrNews, arrTmpListNews
			arrTmpListNews = objListaNews.Items
			Set objTmpCurrNews = arrTmpListNews(FromNews)
			id_news = objTmpCurrNews.getNewsID()
		end if
		
		if bolHasObj then
			Set objCurrentNews = News.findNewsByID(id_news)
			
			response.Write("<span class=""titoloNewsContatti"">"&objCurrentNews.getTitolo() & "</span><br><br>")
			if (Len(objCurrentNews.getAbstract1()) > 0) then response.Write(objCurrentNews.getAbstract1() & "<br>") end if
			if (Len(objCurrentNews.getAbstract2()) > 0) then response.Write(objCurrentNews.getAbstract2() & "<br>") end if
			if (Len(objCurrentNews.getAbstract3()) > 0) then response.Write(objCurrentNews.getAbstract3() & "<br><br>") end if
			response.Write(objCurrentNews.getTesto() & "<br><br>")%>
		
			<br><br>
		<%else
			response.Write("<br/><br/><div align=""center"" class=""titoloNews""><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
		end if%>		
		
		<br><br>		
		<%if (boolMailSent) then
			response.write(lang.getTranslated("frontend.template_form.label.mail_sent"))
		else				
			response.write(lang.getTranslated("frontend.template_form.label.info_"&destination) & "<br><br>")
			
			
			if not(isEmpty(Session("objUtenteLogged"))) then
				Dim objUserLogged, objUserLoggedTmp
				Set objUserLoggedTmp = new UtenteCLass
				Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
				Set objUserLoggedTmp = nothing
				
				if(objUserLogged.getUserActive() = 1) then%>				
				<form action="<%=Application("baseroot") &Application("dir_upload_templ")&"templateFormMail/FormP_dettaglio.asp"%>" method="post" name="form_send_mail">
				  <input type="hidden" name="gerarchia" value="<%=strGerarchia%>">
				  <input type="hidden" name="destination" value="<%=destination%>">
				  <input type="hidden" name="sendmail" value="1">
				  <span class="labelForm"><%=lang.getTranslated("frontend.template_form.label.email")%></span><br>
				  <input type="text" name="email" value="" class="formFieldTXT"><br><br>
				  <span class="labelForm"><%=lang.getTranslated("frontend.template_form.label.oggetto_mail")%></span><br>
				  <input type="text" name="oggetto" value="" class="formFieldTXT"><br><br>
				  <span class="labelForm"><%=lang.getTranslated("frontend.template_form.label.testo_mail")%></span><br>
				  <textarea name="testo" class="formFieldTXTAREA"></textarea><br><br>
				  
				  <a href="javascript:sendMail();"><img src=<%=Application("baseroot")&"/editor/img/t_inserisci.gif"%> vspace="2" hspace="2" border="0" align="middle"></a>		
				</form>				
		<%		end if	
				Set objUserLogged = nothing
			end if
		end if%>
		</td></tr>
		</table>
		<br>
		</td>
		<td class="tdMenuRight">
		<!-- #include virtual="/common/include/MenuContattiDx.inc" -->
		<%if(bolHasObj) then%>
			<!-- #include virtual="/common/include/fileAllegati.inc" -->
			<%Set objCurrentNews = nothing
		end if%>
		</td>
	  </tr>
	</table>
	</td>
  </tr>
  <tr>
    <td class="tdContainerBott">
	<!-- #include virtual="/common/include/bottom.inc" -->
	</td>
  </tr>
</table>
</body>
</html>
<%
Set objListaTargetCat = nothing
Set objListaTargetLang = nothing
Set objListaNews = nothing
Set News = Nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>