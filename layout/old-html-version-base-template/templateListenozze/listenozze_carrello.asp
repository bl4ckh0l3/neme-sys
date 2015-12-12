<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->

<%
Dim News, objListaTargetCat, objListaTargetLang
Dim strGerarchia, strCSS, thisPageName, idCurrentNews, objItem, listOfItem, removeItem, qtaItem

Set News = New NewsClass
strGerarchia = request("gerarchia")
idCurrentNews = request("id_news")
qtaItem = request("qtaItem")
removeItem = request("remove")
thisCatDescDataScad = request("thisCatDescDataScad")
thisPageName = "listenozze_carrello.asp"

if isNull(Session("listenozzeCarrello")) then
	Session("listenozzeCarrello") = ""
end if

if not(isNull(idCurrentNews)) AND not(idCurrentNews = "") then
	Dim isAlreadyAdded, tmpList, tmpSubList, tmpCurrID, tmpCurrQta
		
	if not(isNull(Session("listenozzeCarrello"))) then
		tmpList = Split(Left(Session("listenozzeCarrello"),Len(Session("listenozzeCarrello"))-1), "|", -1, 1)

		if not(isNull(removeItem)) AND not(removeItem = "") then			
			Session("listenozzeCarrello") = ""
			for y=LBound(tmpList) to UBound(tmpList)
				tmpSubList = Split(tmpList(y), "#", -1, 1)
				tmpCurrID = tmpSubList(0)
				tmpCurrQta = tmpSubList(1)
				if not(idCurrentNews = tmpCurrID) then
					Session("listenozzeCarrello") = Session("listenozzeCarrello") & tmpCurrID & "#"&tmpCurrQta&"|"
				end if
			next
		else
			isAlreadyAdded = false
			if not(isNull(tmpList)) AND not(tmpList = "") then			
				Session("listenozzeCarrello") = ""
				for y=LBound(tmpList) to UBound(tmpList)
					tmpSubList = Split(tmpList(y), "#", -1, 1)
					tmpCurrID = tmpSubList(0)
					tmpCurrQta = tmpSubList(1)
					if(idCurrentNews = tmpCurrID) then
						Session("listenozzeCarrello") = Session("listenozzeCarrello") & idCurrentNews & "#"&qtaItem&"|"
						isAlreadyAdded = true
						Exit for
					else
						Session("listenozzeCarrello") = Session("listenozzeCarrello") & tmpCurrID & "#"&tmpCurrQta&"|"						
					end if
				next
			end if
			
			if not(isAlreadyAdded) then
				Session("listenozzeCarrello") = Session("listenozzeCarrello") & idCurrentNews & "#"&qtaItem&"|"
			end if
		end if
	end if
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
function sendMail(){
	if(controllaCampiMail()){
		return true;
		document.form_send_mail.submit();
	}else{
		return false;
	}
}

function controllaCampiMail(){	
	
	if(document.form_send_mail.nome.value == ""){
		alert("<%=lang.getTranslated("frontend.template_pezziunici.js.alert.insert_nome")%>");
		document.form_send_mail.nome.focus();
		return false;
	}	
	
	if(document.form_send_mail.cognome.value == ""){
		alert("<%=lang.getTranslated("frontend.template_pezziunici.js.alert.insert_cognome")%>");
		document.form_send_mail.cognome.focus();
		return false;
	}
	
	var strMail = document.form_send_mail.email.value;
	if(strMail != ""){
		if (strMail.indexOf("@")<2 || strMail.indexOf(".")==-1 || strMail.indexOf(" ")!=-1 || strMail.length<6){
			alert("<%=lang.getTranslated("frontend.template_pezziunici.js.alert.alert.wrong_mail")%>");
			document.form_send_mail.email.focus();
			return false;
		}
	}else if(strMail == ""){
		alert("<%=lang.getTranslated("frontend.template_pezziunici.js.alert.insert_mail")%>");
		document.form_send_mail.email.focus();
		return false;
	}	
	
	if(document.form_send_mail.telefono.value == ""){
		alert("<%=lang.getTranslated("frontend.template_pezziunici.js.alert.insert_telefono")%>");
		document.form_send_mail.telefono.focus();
		return false;
	}		
	
	if(!document.form_send_mail.acceptPrivacy.checked){
		alert("<%=lang.getTranslated("frontend.template_pezziunici.js.alert.confirm_privacy")%>");
		return false;
	}
	var list = document.form_send_mail.listOfItemsInput.value;
	document.form_send_mail.listOfItemsInput.value = list.substring(0, (list.length -2));
	
	//document.form_send_mail.submit();
	return true;
}

function resetForm(){
	document.form_send_mail.reset();
	return false;
}
//-->
</script>
</head>
<body>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
    <td class="tdContainerTopDetail">
	<table cellpadding="0" cellspacing="0" border="0" width="100%">
	<tr>
	<td class="tdHeaderSxCard">	
	<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/"&lang.getTranslated("frontend.template_pezziunici.label.txt_intro")&".gif"%>" align="left" vspace="0" hspace="0" border="0">
	</td>
	<td>	
	<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/logo_header.jpg"%>" align="right" vspace="0" hspace="0" border="0">
	</td>
	</tr>
	</table>
	</td>	
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td class="tdContainerContentCard">	
		<table class="tableContainerInner" border="0" cellpadding="0" cellspacing="0" align="left">
		  <tr>
			<td>
			<%if not(isNull(Session("listenozzeCarrello"))) AND not(Session("listenozzeCarrello") = "") then%>	
				<form action="<%=Application("baseroot") &Application("dir_upload_templ")&"templateListenozze/listenozze_confirm.asp"%>" method="post" name="form_send_mail">
				<table border="0" cellspacing="0" cellpadding="0" align="left">	
				  <tr>
					<td colspan="2"class="txtComposeCard"><%=lang.getTranslated("frontend.template_pezziunici.label.compose_card")%><br><br><br></td>
					<td colspan="5">&nbsp;</td>
				  </tr>		
				  <tr>
					<td colspan="2"><a href="javascript:window.close();"><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/"&lang.getTranslated("frontend.template_pezziunici.label.img_add_to_card")&".gif"%>" align="left" vspace="0" hspace="0" border="0"></a></td>
					<td colspan="5">&nbsp;</td>
				  </tr>		
				  <tr>
					<td colspan="4">&nbsp;</td>
					<td colspan="2" align="right" class="txteuroCard"><%=lang.getTranslated("frontend.template_listenozze.label.txt_dx")%></td>
					<td>&nbsp;</td>
				  </tr>				
				<%listOfItem = Split(Left(Session("listenozzeCarrello"),Len(Session("listenozzeCarrello"))-1), "|", -1, 1)
				Dim objSelNews, iCounter, divAlign, listOfItemsInput, subList, currID, currQta
				iCounter = 1
				listOfItemsInput = ""
				for y=LBound(listOfItem) to UBound(listOfItem)
					subList = Split(listOfItem(y), "#", -1, 1)
					currID = subList(0)
					currQta = subList(1)
					Set objSelNews = News.findNewsByID(currID)			
					listOfItemsInput = listOfItemsInput & "prodotto: " & objSelNews.getTitolo() & ";<br/>codice prodotto: " & objSelNews.getAbstract3() & ";<br/>prezzo: euro&nbsp;" & objSelNews.getTesto() & ";<br/>quantità: " & currQta & ";<br/><br/>||"%>				
              		<tr>
						<td class="txtQta"><%=lang.getTranslated("frontend.template_listenozze.label.qta")%>&nbsp;<%=currQta%>&nbsp;<%if(Cint(currQta) > 1) then response.write(lang.getTranslated("frontend.template_listenozze.carrello.label.qta_pieces")) else response.write(lang.getTranslated("frontend.template_listenozze.carrello.label.qta_piece")) end if%></td>						
						<td class="txtPrenotaCard">Cod.&nbsp;<%=objSelNews.getAbstract3()%></td>
						<td class="attachSmallCard"><!-- #include file="include/fileAllegati.inc" --></td>
						<td class="txtTitolo"><%=objSelNews.getTitolo()%></td>
						<td class="txteuroCard2">Euro</td>
						<td class="txteuroCard"><%=objSelNews.getTesto()%></td>
						<td align="left" valign="bottom" nowrap><a class="txtRemoveFromList" href="<%=Application("baseroot") & "/templates/templateListenozze/"&thisPageName&"?gerarchia="&strGerarchia&"&id_news="&objSelNews.getNewsID()&"&remove=1"%>"><%=lang.getTranslated("frontend.template_pezziunici.label.remove_item")%></a></td>
              		</tr>				
				<%next%>
				  <tr>
				  <td colspan="7" height="40">
				  <input type="hidden" name="listOfItemsInput" value="<%=listOfItemsInput%>">
				  <input type="hidden" name="gerarchia" value="<%=strGerarchia%>">
					<input type="hidden" name="thisCatDescDataScad" value="<%=thisCatDescDataScad%>">			
				  <input type="hidden" name="mailTo" value="<%=Application("mail_receiver")%>">
				  </td>
				  </tr>
				  <!--<tr>
				  <td colspan="2">&nbsp;</td>
				  <td colspan="3" class="formLabelIntro"><%'=lang.getTranslated("frontend.template_pezziunici.label.testo_intro_mail")%><br/>
				  <span class="formLabelIntro2"><%'=lang.getTranslated("frontend.template_pezziunici.label.testo_intro_mail2")%></span><br/>
				  <img src="<%'=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="10" align="top">
				  </td>
				  </tr>-->
				  <tr>
				  <td colspan="2" rowspan="15" class="txtComposeForm"><%=lang.getTranslated("frontend.template_listenozze.label.compile_form")%><br/>
				  <span class="txtComposeFormSmall"><%=lang.getTranslated("frontend.template_listenozze.label.compile_form_small")%></span><br/>
				  <img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/folder_carrello.gif"%>" align="left" vspace="25" hspace="0" border="0"></td>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.nome")%></td>
				  <td colspan="3"><input type="text" name="nome" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.cognome")%></td>
				  <td colspan="3"><input type="text" name="cognome" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.azienda")%></td>
				  <td colspan="3"><input type="text" name="azienda" value="" class="formFieldTXT"></td>
				  </tr>
				  <!--<tr>
				  <td class="formLabel"><%'=lang.getTranslated("frontend.template_pezziunici.label.pivacf")%></td>
				  <td colspan="3"><input type="text" name="pivacf" value="" class="formFieldTXT"></td>
				  </tr>-->
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.indirizzo")%></td>
				  <td colspan="3"><input type="text" name="indirizzo" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.cap_city")%></td>
				  <td colspan="3"><input type="text" name="cap" value="" class="formFieldTXTShort"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="15" width="4" align="absmiddle"><input type="text" name="citta" value="" class="formFieldTXTMedium"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.nazione")%></td>
				  <td colspan="3"><input type="text" name="nazione" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.telefono")%></td>
				  <td colspan="3"><input type="text" name="telefono" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.email")%></td>
				  <td colspan="3"><input type="text" name="email" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.web")%></td>
				  <td colspan="3"><input type="text" name="web" value="" class="formFieldTXT"></td>
				  </tr>
				  <tr>
				  <td colspan="4" class="formLabel">
				  <br/><%=lang.getTranslated("frontend.template_pezziunici.label.info_by")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="infoBy" value="mail" checked><%=lang.getTranslated("frontend.template_pezziunici.label.info_by_email")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="infoBy" value="tel"><%=lang.getTranslated("frontend.template_pezziunici.label.info_by_tel")%>
				  <br/><br/></td>
				  </tr>
				  <tr>
				  <td colspan="4" class="formLabel">
				  <img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="6" align="top"><br>
				  <%=lang.getTranslated("frontend.template_pezziunici.label.testo_mail")%><br/>
				  <textarea name="testo" rows="3" class="formFieldTXTAREA"></textarea></td>
				  </tr>
				  <tr>
				  <td colspan="4" class="formLabel">
				  <img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="15" align="top"><br/>
				  <span class="formLabel"><%=lang.getTranslated("frontend.template_pezziunici.label.info_privacy")%></span><br>
				  <textarea name="testo_privacy" rows="3" class="formFieldTXTAREAPrivacy"><%=lang.getTranslated("frontend.template_contatti.label.info_privacy_law")%></textarea>
				  </td>
				  </tr>
				  <tr>
				  <td colspan="4" class="formLabelSmall">
				  <br><input type="checkbox" name="acceptPrivacy" value="1" hspace="0" vspace="0"><%=lang.getTranslated("frontend.template_pezziunici.label.privacy_accept")%>
				  </td>
				  </tr>
				  <tr>
				  <td colspan="4" align="center">
				  <br><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" vspace="0" hspace="0" border="0" height="8" align="top"><br/>
				  <input type="image" name="reset" onclick="return resetForm();" src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/"&lang.getTranslated("frontend.template_pezziunici.button.cancel.label")&".gif"%>" vspace="0" align="absmiddle">&nbsp;<input type="image" name="submit" onclick="return sendMail();" src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/"&lang.getTranslated("frontend.template_pezziunici.button.send.label")&".gif"%>" vspace="0" align="absmiddle">
				  </td>
				  </tr>
				  <tr>
				  <td colspan="4" align="center">&nbsp;</td>
				  </tr>
				  </table>
				</form>	
			<%else%>
				<div align="center"><b><%=lang.getTranslated("frontend.template_pezziunici.label.empty_card")%></b></div>
			<%end if%>	
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
Set News = Nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>