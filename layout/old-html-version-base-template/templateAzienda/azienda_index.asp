<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->

<%
Dim News, objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim strGerarchia, strCSS, totPages, newsXpage, numPage

Set News = New NewsClass
strGerarchia = request("gerarchia")
order_news_by = 2
newsXpage = 5
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
	<table  class="tableContent" border="0" align="left" cellpadding="0" cellspacing="0">
	  <tr>
		<td class="tdMenu"><br>
		<!-- #include virtual="/common/include/MenuFruizione.inc" --><br>
		<br><br>
		</td>
		<td class="tdContent"><br>
		<!-- #include virtual="/common/include/MenuTips.inc" --><br>
		<hr>
		<%
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
			<table class="tableContentLista" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td class="tdContentLista">
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
						
				objTmpNews = objListaNews.Items%>
				<table border="0" cellspacing="0" cellpadding="0">
				<%for newsCounter = FromNews to ToNews
					Set objSelNews = objTmpNews(newsCounter)%>
					<tr>
					<td align="left" valign="middle" width="80%">
					<br><b><%=objSelNews.getTitolo()%></b><br>
					<%					
					if (Len(objSelNews.getAbstract1()) > 0) then response.Write(objSelNews.getAbstract1() & "<br>") end if
					if (Len(objSelNews.getAbstract2()) > 0) then response.Write(objSelNews.getAbstract2() & "<br>") end if
					if (Len(objSelNews.getAbstract3()) > 0) then%> <a class="titoloNews" href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&id_news="&objSelNews.getNewsID()&"&page="&numPage%>"><%=objSelNews.getAbstract3()%></a><br><br><%end if
					%></td>
					<td align="center" valign="middle" width="20%">&nbsp;
					<%if not(isNull(objSelNews.getFilePerNews())) AND not(isEmpty(objSelNews.getFilePerNews())) then
						Set objListaFilePerNews = objSelNews.getFilePerNews()					
						for each xObjFile in objListaFilePerNews
							Set objFileXNews = objListaFilePerNews(xObjFile)
							iTypeFile = objFileXNews.getFileTypeLabel()
							if(Cint(iTypeFile) = 1) then%>	
								<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" hspace="10" vspace="5" border="0" width="100" height="100" alt="<%=objSelNews.getTitolo()%>" />
								<%Exit for
							end if
							Set objFileXNews = nothing	
						next					
						Set objListaFilePerNews = nothing
					end if
										
					Set objSelNews = nothing%><td>
					</tr>
					<tr>
					<td colspan="2" height="1" align="left" bgcolor="#000000"><img src="<%=Application("baseroot") & "/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0" width="5" height="1" align="top"/><td>
					</tr>
				<%next%>
				</table>				
				</td>
			  </tr>
			  <tr>
				<td class="tdContentPaginazione">
				<%if(totPages > 1) then%>
					<hr align="center" width="100%">
					<%call PaginazioneFrontend(totPages, numPage, strGerarchia, controller_prod_page, "")
				end if%>
				</td>
			  </tr>
			</table>
		<%else
			response.Write("<br/><br/><div align=""center""><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
		end if%>
		</td>
		<td class="tdMenuRight">&nbsp;</td>		
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