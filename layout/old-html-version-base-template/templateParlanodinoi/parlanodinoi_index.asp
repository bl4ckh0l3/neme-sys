<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->

<%
Dim News, objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim strGerarchia, strCSS, totPages, newsXpage, numPage

Set News = New NewsClass
strGerarchia = request("gerarchia")
order_news_by = 12
newsXpage = 10
numPage = 1

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
<SCRIPT language="javascript">
function changeBackgroundImg(id, active, resetCols){
	var element = document.getElementById(id);
	if(active == 1){
		//alert("tdPhotoAttachmentActive");
		if(resetCols = 1){
			var elementTr = document.getElementById("trImgPhoto");
			var cols = elementTr.childNodes
			
			for(var y=0; y<cols.length; y++){	
				var col = cols.item(y);
				//alert("col: " + col.nodeName);
				if(col.getAttribute("class") == "tdPhotoAttachmentActive" || col.getAttribute("className") == "tdPhotoAttachmentActive"){
					//alert("found col");
					col.setAttribute("class", "tdPhotoAttachment"); 
					col.setAttribute("className", "tdPhotoAttachment"); 
				}			
			}
		}
		
		element.setAttribute("class", "tdPhotoAttachmentActive"); 
		element.setAttribute("className", "tdPhotoAttachmentActive"); 
	}else{
		//alert("tdPhotoAttachment");
		element.setAttribute("class", "tdPhotoAttachment"); 
		element.setAttribute("className", "tdPhotoAttachment"); 
	}
}
</script>
</head>
<%
Dim referenceAnchorId
if((not(isNull(request("anchor"))) AND not(request("anchor") = ""))) then
referenceAnchorId = News.findNewsByTitolo(Left(request("anchor"),InStr(1,request("anchor"),"#",1)-1)).getNewsID()
end if
%>
<body onload="MM_preloadImages()<%if((not(isNull(referenceAnchorId)) AND not(referenceAnchorId = ""))) then response.write(";changeBackgroundImg('img"&referenceAnchorId&"0',1,1);") end if%>">
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td class="tdTableContainerSxDx">&nbsp;</td>
    <td class="tdContainerTop"><!-- #include virtual="/common/include/header.inc" --></td>	
  	<td class="tdTableContainerSxDx">&nbsp;</td>
  </tr>
  <tr>
  	<td class="tdTableContainerSxDx">&nbsp;</td>
    <td class="tdContainerContent">
	<table class="tableContent" border="0" align="left" cellpadding="0" cellspacing="0">
	  <tr>
		<td colspan="3">
		<img src="<%=Application("baseroot")&"/templates/templatePhoto/img/header.jpg"%>" vspace="0" hspace="0" border="0" align="top">		
		</td>
		</tr>
	  <tr>
	  	<td colspan="3" class="tdContentColspanMenu"><!-- #include virtual="/common/include/MenuFruizione.inc" --></td>
	  </tr><%
		'************** codice per la lista news e paginazione
		Dim bolHasObj
		bolHasObj = false
		
		on error Resume Next
		if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
			Set objListaNews = News.findNews(null, null, null, objListaTargetCat, objListaTargetLang, null, null, 1, order_news_by)	
			
			if(objListaNews.Count > 0) then		
				bolHasObj = true
			end if
		end if
			
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			bolHasObj = false
		end if			
		
		'if (not(strComp(typename(objListaNews), "ExceptionClass", 1) = 0) _
		'AND isObject(objListaNews)) _
		'AND not(isNull(objListaNews)) _
		'AND not(isEmpty(objListaNews)) then
		
		if(bolHasObj) then%>
		  <tr>
			<td class="tdContentSx"><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templatePhoto/img/pulsante_arancione.gif"%>" vspace="32" hspace="0" border="0" align="top"></td>
			<td class="tdContentTitle">
			<span class="titoloNews"><%=lang.getTranslated("photo_gallery")%></span>	
			</td>
			<td class="tdContentDx">
			<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templatePhoto/img/logo_santenini_small.gif"%>" vspace="0" hspace="0" border="0"><br>
			<img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="20" vspace="0" hspace="0" border="0">
			</td>
		  </tr>
		  <tr>
			<td class="tdContentSx">&nbsp;</td>	  
			<td colspan="2" class="tdContentColspan">		
				<%Dim objSelNews, newsCounter, iIndex, objTmpNews, FromNews, ToNews, Diff
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
						
				objTmpNews = objListaNews.Items		
				for newsCounter = FromNews to ToNews
				'for newsCounter = 0 to objListaNews.Count -1
					Set objSelNews = objTmpNews(newsCounter)%>
					<a name="<%=Replace(objSelNews.getTitolo()," " , "")%>" class="titoloNewsPhoto"><%=objSelNews.getTitolo()%></a><br>
					
					<%if(bolHasObj) then%>
						<!-- #include file="include/fileAllegati.inc" -->
					<%end if%>				
					
					<%Set objSelNews = nothing
				next%>				
				<%if(totPages > 1) then%>
					<hr align="center" width="100%">
					<%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("controller_page"), "")
				end if%>		
		<%else%>

		  <tr>
			<td class="tdContentSx">&nbsp;</td>
			<td class="tdContent">
			<div align="center"><b><%=lang.getTranslated("portal.commons.templates.label.page_in_progress")%></b></div></td>
			<td class="tdContentDx">
			<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templatePhoto/img/logo_santenini_small.gif"%>" vspace="0" hspace="0" border="0"><br>
			<img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="20" vspace="0" hspace="0" border="0">
		<%end if%>	
		</td>
	  </tr>  
	  
	  <tr>
		<td colspan="2">
		<!-- #include virtual="/common/include/MenuFruizioneBottom.inc" -->			
		</td>
		<td class="tdContentDx">&nbsp;</td>
	</tr>
	</table>	</td>
  	<td class="tdTableContainerSxDx">&nbsp;</td>
  </tr>
  <tr>
  	<td class="tdTableContainerSxDx">&nbsp;</td>
    <td class="tdContainerBott">
	<!-- #include virtual="/common/include/bottom.inc" -->
	</td>
  	<td class="tdTableContainerSxDx">&nbsp;</td>
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