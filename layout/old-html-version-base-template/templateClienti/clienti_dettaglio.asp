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
	<table class="tableContent" border="0" align="left" cellpadding="0" cellspacing="0">
	  <tr>
		<td colspan="2" class="tdContentColspan">
		<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateClienti/img/header.jpg"%>" vspace="0" hspace="0" border="0" align="top">		
		</td>
		</tr>
	  <tr>
	  	<td colspan="2" class="tdContentColspanMenu"><!-- #include virtual="/common/include/MenuFruizione.inc" --></td>
	  </tr>
	  <tr>
		<td class="tdContent">
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
			Set objCurrentNews = News.findNewsByID(id_news)%>
			<%if (Len(objCurrentNews.getAbstract1()) > 0) then response.Write(objCurrentNews.getAbstract1() & "<br>") end if
			if (Len(objCurrentNews.getAbstract2()) > 0) then response.Write(objCurrentNews.getAbstract2() & "<br>") end if
			if (Len(objCurrentNews.getAbstract3()) > 0) then response.Write(objCurrentNews.getAbstract3() & "<br><br>") end if
			response.Write(objCurrentNews.getTesto())%>
			
		<%else%>
			<%response.Write("<div align=""center"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
		end if%>
		</td>
		<td class="tdContentDx">
		<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateClienti/img/logo_santenini_small.gif"%>" vspace="0" hspace="0" border="0"><br>
		<img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="20" vspace="0" hspace="0" border="0"><br>
		<%if(bolHasObj) then%>
			<!-- #include file="include/fileAllegati.inc" -->
			<%Set objCurrentNews = nothing
		end if%>
		</td>
	  </tr>
	  <tr>
		<td>
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