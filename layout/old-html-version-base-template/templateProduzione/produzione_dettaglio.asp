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
thisPageName = "produzione_dettaglio.asp"
idCurrentNews = -1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if%>
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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<div id="loadingCatalogo"><img src="<%=Application("baseroot")&"/common/img/loading.gif"%>" vspace="0" hspace="0" border="0" alt="Loading..." width="200" height="50"></div>
<script>
function finish(){
	document.getElementById("loadingCatalogo").style.visibility = "hidden";
	}
</script>
</head>
<body onload="finish()">
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
	<td width="123"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" width="123" height="2" vspace="0" hspace="0" border="0"></td>
	<td class="tdHeaderCenter">
	<%if(bolHasObj) then							
		objTmpNews = objListaNews.Items		
		for newsCounter = FromNews to ToNews
			Set objSelNews = objTmpNews(newsCounter)%>
			<%if not(isNull(objSelNews.getFilePerNews())) AND not(isEmpty(objSelNews.getFilePerNews())) then
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
					Dim hasSmallImg
					hasSmallImg = false
					
					for each xObjFile in objListaFilePerNews
						Set objFileXNews = objListaFilePerNews(xObjFile)					
						
						select case objFileXNews.getFileTypeLabel()
						case 1
							hasSmallImg = true
							Exit for
						case else
						end select
						Set objFileXNews = nothing	
					next
					
					if (cbool(hasSmallImg)) then%>
						<script type="text/javascript" language="JavaScript">
						<!--
						<%' Lista immagini grandi
						for each xObjFile in objListaFilePerNews
							Set objFileXNews = objListaFilePerNews(xObjFile)
							if(objFileXNews.getFileTypeLabel() = 2) then%>	
							Buffer("<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>");


							<%end if
							Set objFileXNews = nothing	
						next%>			
						//-->
						</script>
						<%' Lista immagini piccole
						Dim bigImgID, xObjFileTmp
						for each xObjFile in objListaFilePerNews
							Set objFileXNews = objListaFilePerNews(xObjFile)
							bigImgID = objFileXNews.getFileID()				
							Dim srcFileBig
							if(objFileXNews.getFileTypeLabel() = 1) then						
								srcFileBig = Replace(objFileXNews.getFilePath(),".jpg","_zoom.jpg")%>	
								<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" align="center" hspace="6" vspace="0" border="0" alt="<%=objFileXNews.getFileDida()%>" onmouseover="changeBackgroundImg('showroomRollover','<%=Application("dir_upload_news")&srcFileBig%>');" onmouseout="changeBackgroundImg('showroomRollover','<%=Application("baseroot")&"/templates/templateShowroom/img/base_rollover_big.jpg"%>');">
								<%objListaFilePerNews.remove(xObjFile)
							end if
							Set objFileXNews = nothing	
						next
					end if	
				end if
				Set objListaFilePerNews = nothing
			end if%>	
			<%Set objSelNews = nothing
		next	
	end if%>
	</td>
	<td>	
	<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateProduzione/img/logo_header.jpg"%>" align="right" vspace="0" hspace="0" border="0">
	</td>
	</tr>
	</table>
	</td>	
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td class="tdContainerContent">	
		<table class="tableContainerInner" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
		  	<td class="tdPezziuniciSx">&nbsp;</td>
			<td align="center" height="34" valign="middle" nowrap>	
			<%if(bolHasObj) then					
				if(totPages > 1) then%>
					<%call PaginazioneFrontend(totPages, numPage, strGerarchia, thisPageName, "")
				end if	
			end if%></td>
		  	<td class="tdPezziuniciDx">&nbsp;</td>
		  </tr>
		  <tr>
			<td class="tdPezziuniciSx">
		<%if(bolHasObj) then							
			objTmpNews = objListaNews.Items		
			for newsCounter = FromNews to ToNews
			'for newsCounter = 0 to objListaNews.Count -1
				Set objSelNews = objTmpNews(newsCounter)
				idCurrentNews = objSelNews.getNewsID()%>
				<%=objSelNews.getTitolo()%><br><br><br>
				<%=objSelNews.getTesto()%>
				<%Set objSelNews = nothing
			next
		end if%>
		<br><br><br><a href="javascript:window.close();"><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateProduzione/img/"&lang.getTranslated("frontend.template_pezziunici.img_back_catalogo")&".gif"%>" align="left" vspace="0" hspace="0" border="0"></a>
			</td>
			<td class="tdPezziuniciCenter">	
		<%if(bolHasObj) then							
			objTmpNews = objListaNews.Items		
			for newsCounter = FromNews to ToNews
				Set objSelNews = objTmpNews(newsCounter)%>
					<!-- #include file="include/fileAllegati.inc" -->				
				<%Set objSelNews = nothing
			next%>	
		<%else%>
			<div align="center"><b><%=lang.getTranslated("portal.commons.templates.label.page_in_progress")%></b></div>
		<%end if%>	
		</td>
			<td class="tdPezziuniciDx">
			<span class="txtRequestPrice"><%=lang.getTranslated("frontend.template_pezziunici.label.richiesta_prezzi_info")%></span><br/>
			<span class="txtClickToAdd"><%=lang.getTranslated("frontend.template_pezziunici.label.click_to_add")%></span><br/><br/>
			<a href="javascript:openWin('<%=Application("baseroot") &Application("dir_upload_templ")&"templateProduzione/produzione_carrello.asp?gerarchia="&strGerarchia&"&id_news="&idCurrentNews%>','pezziunicicarrello',970,600,150,60)"><img src="<%=Application("baseroot")&"/templates/templatePezziunici/img/carrello.jpg"%>" align="absmiddle" vspace="0" hspace="0" border="0"></a>
			<span class="txtAddToCard"><%=lang.getTranslated("frontend.template_pezziunici.label.add_to_card")%></span>
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