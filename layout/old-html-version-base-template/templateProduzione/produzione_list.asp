<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->

<%
Dim id_news, News, objCurrentNews, strGerarchia, strCSS, objFileXNews, objListaFilePerNews
Dim objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim totPages, newsXpage, numPage

Set News = New NewsClass
id_news = request("id_news")
strGerarchia = request("gerarchia")
order_news_by = 11
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

//-->
</script>
</head>
<body>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
    <td class="tdContainerContentList">	
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
			bolHasObj = false
		end if			
		'response.write("obi:"&typename(objListaNews))
		'response.write("obi:"&objListaNews.Count)
		if(bolHasObj) then
				Dim objSelNews, newsCounter, iIndex, objTmpNews, FromNews, ToNews, Diff, iCounter
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


				'response.write("<b>"&lang.getTranslated("portal.template_parlanodinoi.label.testate")&"</b><br/><br/><br/>")

				'objTmpNews = objListaNews.Items		
				'for newsCounter = FromNews to ToNews
				iCounter = 0
				for each newsCounter in objListaNews
					Set objSelNews = objListaNews(newsCounter)
					'response.write("obi:"&typename(objSelNews))
					if not(isNull(objSelNews.getFilePerNews())) AND not(isEmpty(objSelNews.getFilePerNews())) AND isObject(objSelNews.getFilePerNews()) then
						Set objListaFilePerNews = objSelNews.getFilePerNews()
						'response.write("obi:"&typename(objListaFilePerNews))
						'response.end
						if not(isEmpty(objListaFilePerNews)) then
							
							for each xObjFile in objListaFilePerNews
								Set objFileXNews = objListaFilePerNews(xObjFile)					
								if(objFileXNews.getFileTypeLabel() = 6) then
									iCounter = iCounter +1%><a href="javascript:openWin('<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&id_news="&objSelNews.getNewsID()&"&page="&iCounter%>','pezziunici',screenW,screenH,0,0)"><img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" hspace="5" vspace="0" border="0" alt="<%=objFileXNews.getFileDida()%>"></a>
									<%'objListaFilePerNews.remove(xObjFile)																		
									Exit For
								end if
								Set objFileXNews = nothing
							next
						
						end if
						Set objListaFilePerNews = nothing
					end if
					
					if(iCounter Mod 4 = 0) then%><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="10" vspace="0" hspace="0" border="0" align="top"><br/><%end if

					Set objSelNews = nothing
				next%>
				<%'if(totPages > 1) then%>
					<!--hr align="center" width="100%"-->
					<%'call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("controller_page"), "")
				'end if%>
		<%else
			response.Write("<br/><br/><div align=""center""><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
		end if%>

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