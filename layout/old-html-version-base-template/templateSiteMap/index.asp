<%'On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->

<%
Dim id_news, News, objCurrentNews, strGerarchia, strCSS, objFileXNews, objListaFilePerNews
Dim objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim totPages, newsXpage, numPage

strGerarchia = request("gerarchia")

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if%>
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
		<td colspan="2" class="tdContentColspan">
		<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateSiteMap/img/header.jpg"%>" vspace="0" hspace="0" border="0" align="top">		
		</td>
		</tr>
	  <tr>
		<td colspan="2" class="tdContentColspanMenu"><!-- #include virtual="/common/include/MenuFruizione.inc" --></td>
		</tr>
	  <tr>
		<td class="tdContentMap">
		<%
		Dim menuFruizioneMap, menuCompleteMap, categoriaClassTmpMap
		Dim levelMap, iWidthMap, strSubTmpGerMap, strSubTmpGerMapFiltered
		Dim iLenGerMap, iGerlevelMap, iGerDiffMap, hrefGerMap

		Set menuFruizioneMap = new menuFruizioneClass
		Set menuCompleteMap = menuFruizioneMap.getCompleteMenu()
		Set categoriaClassTmpMap = new CategoriaClass

		if(isNull(strGerarchia) OR strGerarchia = "") then strGerarchia = "00"

		iGerlevelMap = menuFruizioneMap.getLivello(strGerarchia)%>
		<%
		Dim xMap, objCategoriaCheckMap, strHrefMap, menuMapCatDescTrans, menuMapCatDescTextTrans
		for each xMap in menuCompleteMap
			levelMap = menuFruizioneMap.getLivello(xMap)
			iGerDiffMap = levelMap - iGerlevelMap
			menuMapCatDescTrans = "frontend.menu.label."&menuCompleteMap(xMap)
			menuMapCatDescTextTrans = "frontend.menu.label.desc."&menuCompleteMap(xMap)
				
			if(levelMap > 1) then
				iWidthMap = (levelMap-1) * 10 
				iLenGerMap = (levelMap * 2) + (levelMap -1)
				strSubTmpGerMap = Left(strGerarchia, iLenGerMap)
				strSubTmpGerMapFiltered = Left(strSubTmpGerMap, iLenGerMap-3)
				
				if(iGerDiffMap <= 1) then
					if (InStr(1, Left(xMap, iLenGerMap-3), strSubTmpGerMapFiltered, 1) > 0) then
						hrefGerMap = xMap
						strHrefMap = Application("baseroot") & "/common/include/Controller.asp?gerarchia="&hrefGerMap				
						
						'*** Controllo se la categoria contiene news, altrimenti cerco la prima sottocategoria che contenga news
						'*** e imposto la nuova gerarchia come parametro nel link
						Set objCategoriaCheckMap = categoriaClassTmpMap.findCategoriaByGerarchia(hrefGerMap)
						if not(objCategoriaCheckMap.contieneNews()) AND not(objCategoriaCheckMap.contieneProd()) then
							foundGer = categoriaClassTmpMap.findFirstSubCategoriaWithNews(hrefGerMap)
							
							if(isNull(foundGer)) then
								foundGer = categoriaClassTmpMap.findFirstSubCategoriaWithProd(hrefGerMap)
							end if
							
							if not(isNull(foundGer)) then
								strHrefMap = Application("baseroot") & "/common/include/Controller.asp?gerarchia="&foundGer
							else
								strHrefMap = "#"
							end if
						end if
						Set objCategoriaCheckMap = nothing				
						%>
						<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateSiteMap/img/freccina_sottotitolo.gif"%>" vspace="0" hspace="5" border="0" align="absmiddle"><a class="MenuFruizioneMap<%if(strComp(xMap, strSubTmpGerMap, 1) = 0) then response.Write("Active")%>" href="<%=strHrefMap%>"><%if not(isNull(lang.getTranslated(menuMapCatDescTrans))) AND not(lang.getTranslated(menuMapCatDescTrans) = "") then response.write(lang.getTranslated(menuMapCatDescTrans)) else response.Write(menuCompleteMap(xMap)) end if%></a>
						<%if not(isNull(lang.getTranslated(menuMapCatDescTextTrans))) AND not(lang.getTranslated(menuMapCatDescTextTrans) = "") then response.write("<div class=""txtContentDescMap"">"&lang.getTranslated(menuMapCatDescTextTrans)&"</div>") else response.write("<br/>") end if%>
						<br/>

					<%end if
				end if
			else
				iWidthMap = 0
				iLenGerMap = 2
				strSubTmpGerMap = Left(strGerarchia, iLenGerMap)
				hrefGerMap = xMap
				strHrefMap = Application("baseroot") & "/common/include/Controller.asp?gerarchia="&hrefGerMap
				
				'*** Controllo se la categoria contiene news, altrimenti cerco la prima sottocategoria che contenga news
				'*** e imposto la nuova gerarchia come parametro nel link
				Set objCategoriaCheckMap = categoriaClassTmpMap.findCategoriaByGerarchia(hrefGerMap)
				if not(objCategoriaCheckMap.contieneNews()) AND not(objCategoriaCheckMap.contieneProd()) then
					foundGer = categoriaClassTmpMap.findFirstSubCategoriaWithNews(hrefGerMap)
					
					if(isNull(foundGer)) then
						foundGer = categoriaClassTmpMap.findFirstSubCategoriaWithProd(hrefGerMap)
					end if
					
					if not(isNull(foundGer)) then
						strHrefMap = Application("baseroot") & "/common/include/Controller.asp?gerarchia="&foundGer
					else
						strHrefMap = "#"
					end if
				end if
				Set objCategoriaCheckMap = nothing%>
				<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateSiteMap/img/freccina_sottotitolo.gif"%>" vspace="0" hspace="5" border="0" align="absmiddle"><a class="MenuFruizioneMap<%if(strComp(xMap, strSubTmpGerMap, 1) = 0) then response.Write("Active")%>" href="<%=strHrefMap%>"><%if not(isNull(lang.getTranslated(menuMapCatDescTrans))) AND not(lang.getTranslated(menuMapCatDescTrans) = "") then response.write(lang.getTranslated(menuMapCatDescTrans)) else response.Write(menuCompleteMap(xMap)) end if%></a>
				<%if not(isNull(lang.getTranslated(menuMapCatDescTextTrans))) AND not(lang.getTranslated(menuMapCatDescTextTrans) = "") then response.write("<div class=""txtContentDescMap"">"&lang.getTranslated(menuMapCatDescTextTrans)&"</div>") else response.write("<br/>") end if%>
				<br/>
			<%end if	
		next
		Set categoriaClassTmpMap = nothing
		Set menuCompleteMap = nothing
		Set menuFruizioneMap = nothing
		%>		
		</td>
		<td class="tdContentDxMap">
		<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateSiteMap/img/logo_santenini_small.gif"%>" vspace="0" hspace="0" border="0"><br>
		<img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="20" vspace="0" hspace="0" border="0"><br>
		</td>
	  </tr>
	  <tr>
		<td class="tdContentColspanMap"><!-- #include virtual="/common/include/MenuFruizioneBottom.inc" --><br/></td>
		<td class="tdContentDx">&nbsp;</td>
	</tr>
	</table>
	</td>
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