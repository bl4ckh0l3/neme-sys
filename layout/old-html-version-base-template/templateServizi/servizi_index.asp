<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->

<%
Dim News, objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim strGerarchia, strCSS, totPages, newsXpage, numPage

Set News = New NewsClass
strGerarchia = request("gerarchia")
order_news_by = 1
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
    <td class="tdContainerTop">
	<!-- #include virtual="/common/include/header.inc" -->	
	</td>
  </tr>
  <tr>
    <td class="tdContainerContent">
	<table  class="tableContent" border="0" align="left" cellpadding="0" cellspacing="0">
	  <tr>
		<td class="tdMenu"><!-- #include virtual="/common/include/MenuFruizione.inc" --></td>
		<td class="tdContent"><!-- #include virtual="/common/include/MenuTips.inc" -->
		<p class="txtIntroServizi"><%=lang.getTranslated("frontend.template_servizi.label.intro_servizi")%></p>
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
			<table class="tableContentService" border="0" align="center" cellspacing="0" cellpadding="0">
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
					  <tr>
						<td>
						<%if (not(isNull(objSelNews.getFilePerNews())) AND isObject(objSelNews.getFilePerNews()) AND not(isEmpty(objSelNews.getFilePerNews()))) then
							Set objListaFilePerNews = objSelNews.getFilePerNews()					
							for each xObjFile in objListaFilePerNews
								Set objFileXNews = objListaFilePerNews(xObjFile)
								iTypeFile = objFileXNews.getFileTypeLabel()
								if(Cint(iTypeFile) = 1) then%>	
									<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" hspace="0" vspace="0" border="0" alt="<%=objSelNews.getTitolo()%>" />
									<%Exit for
								end if
								Set objFileXNews = nothing	
							next					
							Set objListaFilePerNews = nothing
						end if%>						
						<p><a class="titoloNews" href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&id_news="&objSelNews.getNewsID()&"&page="&numPage%>"><%=objSelNews.getTitolo()%></a><br/>	
						<%if (Len(objSelNews.getAbstract1()) > 0) then response.Write(objSelNews.getAbstract1()) end if%></p>				
						</td>
					  </tr>
					<%Set objSelNews = nothing
				next%>
			</table>
			<%if(totPages > 1) then%>
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td  class="tdContentPaginazione"><p>
					<%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("controller_page"), "")%>
					</p>
				</td>
			</tr>
			</table>
			<%end if%>
		<%else
			response.Write("<br/><br/><div align=""center""><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
		end if%>
		</td>
		<td class="tdMenuRight"><!-- #include virtual="/common/include/MenuContattiDx.inc" --></td>		
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