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
<script type="text/javascript" language="JavaScript">
<!--
var bufferImage = new Array();

function Buffer(filename) {
    var i = bufferImage.length;
    bufferImage[i] = new Image();
    bufferImage[i].src = filename;
}

function changeBackgroundImg(id, img){
	var element = document.getElementById(id);
	element.setAttribute("src", img); 
}

function sendMail(){
	if(controllaCampiMail()){
		return true;
		//document.form_send_mail.submit();
	}else{
		return false;
	}
}

function controllaCampiMail(){	
	
	if(document.form_send_mail.nome.value == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_nome")%>");
		document.form_send_mail.nome.focus();
		return false;
	}	
	
	if(document.form_send_mail.cognome.value == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_cognome")%>");
		document.form_send_mail.cognome.focus();
		return false;
	}
	
	var strMail = document.form_send_mail.email.value;
	if(strMail != ""){
		if (strMail.indexOf("@")<2 || strMail.indexOf(".")==-1 || strMail.indexOf(" ")!=-1 || strMail.length<6){
			alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.alert.wrong_mail")%>");
			document.form_send_mail.email.focus();
			return false;
		}
	}else if(strMail == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_mail")%>");
		document.form_send_mail.email.focus();
		return false;
	}	
	
	if(document.form_send_mail.telefono.value == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_telefono")%>");
		document.form_send_mail.telefono.focus();
		return false;
	}		
	
	if(!document.form_send_mail.acceptPrivacy.checked){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.confirm_privacy")%>");
		return false;
	}
	
	return true;
}
//-->
</script>

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
    <td class="tdContainerContent">
		<table border="0" cellpadding="0" cellspacing="0" align="left">
		  <tr>
			<td class="contatti2Left">
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
					response.Write(objCurrentNews.getTesto())
					Set objCurrentNews = nothing
				else
					response.Write("<br/><br/><div align=""center""><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
				end if%>			
			</td>
			<td><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateContatti2/img/sfondo_center.jpg"%>" vspace="0" hspace="0" border="0"></td>
			<td class="contatti2Right">							
				<form action="<%=Application("baseroot") &Application("dir_upload_templ")&"templateContatti2/contatti2_confirm.asp"%>" method="post" name="form_send_mail" onSubmit="return sendMail();">
				  <input type="hidden" name="gerarchia" value="<%=strGerarchia%>">
				  <input type="hidden" name="mailTo" value="<%=Application("mail_receiver")%>">
				  <table border="0" cellpadding="0" cellspacing="0" align="top">
				  <tr>
				  <td colspan="2" class="formLabelIntro"><%=lang.getTranslated("frontend.template_contatti.label.testo_intro_mail")%><br/>
				  <span class="formLabelIntro2"><%=lang.getTranslated("frontend.template_contatti.label.testo_intro_mail2")%></span><br/>
				  <img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="5" align="top"></td></tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_contatti.label.nome")%></td>
				  <td><input type="text" name="nome" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_contatti.label.cognome")%></td>
				  <td><input type="text" name="cognome" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_contatti.label.email")%></td>
				  <td><input type="text" name="email" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_contatti.label.telefono")%></td>
				  <td><input type="text" name="telefono" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_contatti.label.indirizzo")%></td>
				  <td><input type="text" name="indirizzo" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_contatti.label.cap_city")%></td>
				  <td><input type="text" name="cap" value="" class="formFieldTXTShort">&nbsp;<input type="text" name="citta" value="" class="formFieldTXTMedium"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_contatti.label.nazione")%></td>
				  <td><input type="text" name="nazione" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td colspan="2" class="formLabel">
				  <img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="6" align="top"><br>
				  <%=lang.getTranslated("frontend.template_contatti.label.testo_mail")%><br/>
				  <textarea name="testo" rows="3" class="formFieldTXTAREA"></textarea></td></tr>
				  <tr>
				  <td colspan="2" class="formLabel">
				  <%=lang.getTranslated("frontend.template_contatti.label.info_by")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="infoBy" value="mail" checked><%=lang.getTranslated("frontend.template_contatti.label.info_by_email")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="infoBy" value="tel"><%=lang.getTranslated("frontend.template_contatti.label.info_by_tel")%>
				  </td></tr>
				  <tr>
				  <td colspan="2" class="formLabel">
				  <img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="4" align="top"><br/>
				  <span class="formLabel"><%=lang.getTranslated("frontend.template_contatti.label.info_privacy")%></span><br>
				  <textarea name="testo_privacy" rows="3" class="formFieldTXTAREAPrivacy"><%=lang.getTranslated("frontend.template_contatti.label.info_privacy_law")%></textarea>
				  </td></tr>
				  <tr>
				  <td colspan="2" class="formLabelSmall">
				  <input type="checkbox" name="acceptPrivacy" value="1" hspace="0" vspace="0"><%=lang.getTranslated("frontend.template_contatti.label.privacy_accept")%>
				  </td></tr>
				  <td colspan="2" align="center">
				  <img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="8" align="top"><br/>
				  <input type="reset" name="reset" value="<%=lang.getTranslated("frontend.template_contatti.button.cancel.label")%>" class="formFieldButton" vspace="0" align="absmiddle">&nbsp;<input type="submit" name="submit" value="<%=lang.getTranslated("frontend.template_contatti.button.send.label")%>" class="formFieldButton" vspace="0" align="absmiddle">
				  </td></tr>
				  </table>
				</form>		
			</td>
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
	<td align="left" valign="top" width="194"><!--<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateContatti2/img/bottom_title_"&lang.getLangCode()&".gif"%>" vspace="0" hspace="0" border="0">--></td>
	<td align="left" valign="top">
	<img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="20" align="top"><br/>
	<span class="bottCopyright"><%=lang.getTranslated("frontend.template_contatti.label.testo_bottom_copyright")%></span><br/>
	<img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="10" align="top"><br/>
	<span class="bottOrari"><%=lang.getTranslated("frontend.template_contatti.label.testo_bottom_orari")%></span>
	</td>
	</tr>
	</table>
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