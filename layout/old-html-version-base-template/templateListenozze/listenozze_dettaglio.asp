<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->

<%
Dim News, objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim strGerarchia, strCSS, totPages, newsXpage, numPage, thisPageName, idCurrentNews

Dim stato
stato = 1
if(request("is_preview_content")) then
	stato = null
end if

Set News = New NewsClass
strGerarchia = request("gerarchia")
order_news_by = 11
newsXpage = 1
numPage = 1
thisPageName = "listenozze_dettaglio.asp"
idCurrentNews = -1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if%>
<!-- #include virtual="/common/include/setTemplateTargetList.inc" -->
<%
Dim categoriaClassTmp, objTmp, menuCompleteCatDescTrans, thisCatDesc, thisCatDescDataScad
thisCatDesc = ""
menuCompleteCatDescTrans = "frontend.menu.label."
Set categoriaClassTmp = new CategoriaClass
Set objTmp = categoriaClassTmp.findCategoriaByGerarchia(strGerarchia)
if not(isNull(objTmp)) AND not(isEmpty(objTmp)) then
	thisCatDesc = objTmp.getCatDescrizione()
	thisCatDescDataScad = thisCatDesc
	menuCompleteCatDescTrans = menuCompleteCatDescTrans & thisCatDesc
	if not(isNull(lang.getTranslated(menuCompleteCatDescTrans & thisCatDesc))) AND not(lang.getTranslated(menuCompleteCatDescTrans & thisCatDesc) = "") then
		thisCatDesc = lang.getTranslated(menuCompleteCatDescTrans & thisCatDesc)
	end if
	Set objTmp = nothing
end if
Set categoriaClassTmp = nothing
%>
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
var screenW = 640, screenH = 480;
if (parseInt(navigator.appVersion)>3) {
 screenW = screen.width-20;
 screenH = screen.height-20;
}
else if (navigator.appName == "Netscape" 
    && parseInt(navigator.appVersion)==3
    && navigator.javaEnabled()
   ) 
{
 var jToolkit = java.awt.Toolkit.getDefaultToolkit();
 var jScreenSize = jToolkit.getScreenSize();
 screenW = jScreenSize.width-20;
 screenH = jScreenSize.height-20;
}


function sendCarrello(theform,reloadCard){
	var thisForm = theform;	
	var url = '<%=Application("baseroot") &Application("dir_upload_templ")&"templateListenozze/listenozze_carrello.asp?"%>';
	
	if(reloadCard){
		url = url +'gerarchia='+thisForm;
		openWin(url,'listenozzecarrello',970,600,150,60);
	}else{
		var gerarchia = thisForm.gerarchia.value;
		var id_news = thisForm.id_news.value;
		var cod_num = thisForm.cod_num.value;
		var price = thisForm.price.value;
		var i;
		var qtaItem = 0;
		if(theform.qta != null){
			if(theform.qta.length == null){
				if(theform.qta.checked){
					qtaItem = theform.qta.value;
				}
			}else{
				for(i=0; i<theform.qta.length; i++){
					if(theform.qta[i].checked){	
						++qtaItem;
					}
				}
			}
		}
		/*if(qtaItem.charAt(strFiles.length -1) == "|"){
			qtaItem = qtaItem.substring(0, qtaItem.length -1);
		}	
		document.form_inserisci.ListFileDaEliminare.value = qtaItem;*/
		
		/*alert("gerarchia: " + gerarchia);
		alert("id_news: " + id_news);
		alert("cod_num: " + cod_num);
		alert("price: " + price);
		alert("qtaItem: " + qtaItem);*/
	
		url = url +'gerarchia='+gerarchia;
		url = url +'&id_news='+id_news;
		//url = url +'&cod_num='+cod_num;
		//url = url +'&price='+price;
		url = url +'&qtaItem='+qtaItem;
		
		if(qtaItem > 0){
			openWin(url,'listenozzecarrello',970,600,150,60);
		}else{
			alert("<%=lang.getTranslated("frontend.template_listenozze.js.label.select_item")%>");
		}
	}
}
//-->
</script>
</head>
<body>
<%
'************** codice per la lista news e paginazione
Dim bolHasObj
bolHasObj = false

on error Resume Next
if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
	Set objListaNews = News.findNews(null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_news_by)	
	
	if(objListaNews.Count > 0) then		
		bolHasObj = true
	end if
end if
	
if Err.number <> 0 then
	bolHasObj = false
end if		

if(bolHasObj) then
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
end if		
%>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
    <td class="tdContainerTopDetail">
	<table cellpadding="0" cellspacing="0" border="0" width="100%">
	<tr>	
	<td width="60"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" width="60" height="2" vspace="0" hspace="0" border="0"></td>
	<td class="tdHeaderCenter">
	<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/"&lang.getTranslated("frontend.template_listenozze.label.txt_intro")&".gif"%>" align="absmiddle" vspace="0" hspace="0" border="0"><%=thisCatDesc%>
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
    <td class="tdContainerContentDett">	
		<table border="0" cellspacing="0" cellpadding="0" class="tableContainerInner">
		  <tr>
			<td width="255">&nbsp;</td>
			<td width="572">
			<table border="0" cellspacing="0" cellpadding="0" width="100%" class="tdScadLista">
              <tr>
                <td width="1%"><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/bordino_top_sx.gif"%>" align="left" vspace="0" hspace="0" border="0"></td>
                <td  width="98%" bgcolor="#C8C7C6" align="center"><%=lang.getTranslated("frontend.template_listenozze.label.scad_lista")%>&nbsp;&nbsp;<%=lang.getTranslated("frontend.template_listenozze.label.data_scad_lista."&thisCatDescDataScad)%></td>
                <td width="1%"><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/bordino_top_dx.gif"%>" align="left" vspace="0" hspace="0" border="0"></td>
              </tr>
            </table></td>
		  </tr>
		  <tr>
			<td width="255" align="left"><br/><%=lang.getTranslated("frontend.template_listenozze.label.txt_sx")%></td>
			<td width="572" align="right"><a href="javascript:sendCarrello('<%=strGerarchia%>',true);"><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/ico_carrello_go.gif"%>" align="absmiddle" vspace="0" hspace="0" border="0" alt="<%=lang.getTranslated("frontend.template_listenozze.label.open_card")%>"></a></td>
		  </tr>
		  <tr>
			<td colspan="2" align="right" valign="top">
			<%if(bolHasObj) then%>			
				<table border="0" cellspacing="0" cellpadding="0">	
				  <tr>
					<td colspan="6">&nbsp;</td>
					<td colspan="2" align="right" class="txtEuro"><%=lang.getTranslated("frontend.template_listenozze.label.txt_dx")%></td>
				  </tr>			
				<%'objTmpNews = objListaNews.Items
				
				'for newsCounter = FromNews to ToNews
				for each newsCounter in objListaNews
					'Set objSelNews = objTmpNews(newsCounter)
					Set objSelNews = objListaNews(newsCounter)
					idCurrentNews = objSelNews.getNewsID()%>
					<form method="post" name="formQta<%=newsCounter%>">
					<input type="hidden" name="gerarchia" value="<%=strGerarchia%>">
					<input type="hidden" name="thisCatDescDataScad" value="<%=thisCatDescDataScad%>">					
					<input type="hidden" name="id_news" value="<%=idCurrentNews%>">
					<input type="hidden" name="cod_num" value="<%=objSelNews.getAbstract3()%>">
					<input type="hidden" name="price" value="<%=objSelNews.getTesto()%>">
              		<tr>
						<td class="txtCheckbox">
						<%if not(objSelNews.getAbstract1() = "") then
							for x = 1 to Cint(objSelNews.getAbstract1())%>
								<div style="float:left;padding-top:0px;padding-bottom:0px;padding-right:5px;padding-left:5px;text-align:center;">
								<%=objSelNews.getAbstract2()%>&nbsp;<%if(Cint(objSelNews.getAbstract2()) > 1) then response.write(lang.getTranslated("frontend.template_listenozze.label.qta_pieces")) else response.write(lang.getTranslated("frontend.template_listenozze.label.qta_piece")) end if%><br/>
								<input type="checkbox" name="qta" value="1"></div>								
							<%Next
						end if%>
						</td>						
						<td class="txtCarrello">
						<%if not(objSelNews.getAbstract1() = "") then 
							if (Cint(objSelNews.getAbstract1()) > 0) then%><a href="javascript:sendCarrello(document.formQta<%=newsCounter%>,false);"><img src="<%=Application("baseroot")&"/templates/templateListenozze/img/ico_carrello.gif"%>" align="absmiddle" vspace="0" hspace="0" border="0"></a><%else%>&nbsp;<%end if 
						end if%></td>
						<td class="txtPrenota"><%if not(objSelNews.getAbstract1() = "") then
						if (Cint(objSelNews.getAbstract1()) > 0) then%><b><%=lang.getTranslated("frontend.template_listenozze.label.txt_prenota")%></b><br/><%end if
						end if%>
						Cod.&nbsp;<%=objSelNews.getAbstract3()%></td>
						<td class="tdAttachments"><!-- #include file="include/fileAllegati.inc" --></td>
						<td class="txtTitolo"><%=objSelNews.getTitolo()%></td>
						<td align="left" valign="bottom">
						<%
						if not(isNull(objSelNews.getFilePerNews())) AND not(isEmpty(objSelNews.getFilePerNews())) then
							Set objListaFilePerNews = objSelNews.getFilePerNews()
							
							if not(isEmpty(objListaFilePerNews)) then
								' LEGENDA TIPI FILE
								'1 = img small
								'2 = img big
								'3 = pdf
								'4 = audio-video
								'5 = others%>				
								<%
								' Lista label tipi file
								Dim hasCardImg
								hasCardImg = false
								
								for each xObjFile in objListaFilePerNews
									Set objFileXNews = objListaFilePerNews(xObjFile)					
									
									select case objFileXNews.getFileTypeLabel()
									case 2
										hasCardImg = true
										Exit for
									case else
									end select
									Set objFileXNews = nothing	
								next
								
								if (cbool(hasCardImg)) then%>
									<%for each xObjFile in objListaFilePerNews
										Set objFileXNews = objListaFilePerNews(xObjFile)				
										if(objFileXNews.getFileTypeLabel() = 2) then%>											
											<a href="javascript:openWin('<%=Application("baseroot") & "/common/include/popup.asp?id_allegato="&objFileXNews.getFileID()%>','listenozzezoom',420,585,150,60)"><img src="<%=Application("baseroot")&"/templates/templateListenozze/img/"&lang.getTranslated("frontend.template_listenozze.label.zoom_img")&".gif"%>" align="absbottom" vspace="0" hspace="0" border="0"></a>
											<%Exit for
										end if
										Set objFileXNews = nothing	
									next								
								end if
							end if
							Set objListaFilePerNews = nothing
						end if						
						%></td>
						<td class="txtEuro">Euro</td>
						<td class="txtEuro"><%=objSelNews.getTesto()%></td>
              		</tr>						
					</form>
					<%Set objSelNews = nothing
				next%>	
				  <tr>
					<td colspan="8">&nbsp;</td>
				  </tr>
            	</table>
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
Set objListaNews = nothing
Set News = Nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>