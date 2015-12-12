<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->

<%
Dim News, objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim strGerarchia, strCSS, totPages, newsXpage, numPage

Set News = New NewsClass
strGerarchia = request("gerarchia")
order_news_by = 12
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
<meta content="The best guide to visit and enjoy Lisbon: restaurants, fado, discos, bars, shopping, hotels, hostels, apartments, tours, rental cars, beaches, sights and others tourism tips for your travel." name=description>

<META name="description" CONTENT="Tourism guide of Lisbon">
<META name="keywords" CONTENT="Lisbon, tourism, guide, tips, portugal, nightlife, fado, restaurants, discos, bars, pubs, shopping, hotels, hostels, apartments, tours, rental cars, arts, sport  beaches, sights">
<META name="owner" CONTENT="Tips Guide Lisboa - www.tipsguidelisboa.com">
<META name="robots" CONTENT="Index, Follow">
<meta name="verify-v1" content="xWdjZBidrZnB2AjIR9jW2fh6A2wSZcMggJ2bQIDDi7I=" />
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
	<table  class="tableContent" border="0" cellpadding="0" cellspacing="0">
	  <tr>
		<td class="tdMenu"><!-- #include virtual="/common/include/MenuFruizione.inc" --></td>
		<td class="tdContent">
		<!-- #include virtual="/common/include/MenuTips.inc" -->
		<%
		'************** codice per la lista news e paginazione
		Dim bolHasObj
		bolHasObj = false
		
		'on error Resume Next
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
				<%for newsCounter = FromNews to ToNews
					Set objSelNews = objTmpNews(newsCounter)%>
					
				  <tr>
				    <td><br/><p>
					<span class="titoloPartner"><%=objSelNews.getTitolo()%></span><br />
					<br />
				    
					<%					
					if (Len(objSelNews.getAbstract1()) > 0) then %>
						<span class="testoInfo"><%=objSelNews.getAbstract1()%></span><br /><br />
					<% end if
					if (Len(objSelNews.getAbstract2()) > 0) then %>
						<a class="titoloNews" href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&id_news="&objSelNews.getNewsID()&"&page="&numPage%>"><%=objSelNews.getAbstract2()%></a><br/><br/>
					<%end if
					if (Len(objSelNews.getAbstract3()) > 0) then%> 
						<%=objSelNews.getAbstract3()%>
					<%end if%>					
					</p>						
					<%if (Len(objSelNews.getAbstract2()) > 0) then %><a href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&id_news="&objSelNews.getNewsID()&"&page="&numPage%>"><img src="<%=Application("baseroot")&"/common/img/bot_infopicture.jpg"%>" alt="<%=objSelNews.getAbstract2()%>" hspace="0" vspace="0" border="0" /></a><%end if%><a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popupInsertNewsComments.asp?id_news="&objSelNews.getNewsID()%>','popupallegati',400,400,100,100);" title="<%=lang.getTranslated("portal.templates.commons.label.see_comments_news")%>"><img src="<%=Application("baseroot")&"/common/img/bot_commenta.jpg"%>" hspace="0" vspace="0" border="0"></a><!--<a href="<%'=Application("baseroot") & "/common/include/Controller.asp?gerarchia=-5&destination=" & Application("dest_speciali")%>"><img src="<%'=Application("baseroot")&"/common/img/bot_mail.jpg"%>" alt="<%'=lang.getTranslated("cat_contact_special")%>" hspace="0" vspace="0" border="0" /></a>--></td>
				    <td align="center" valign="top">&nbsp;						
					<%
					'response.write("statistiche:<br>")
					'response.write(typename(objSelNews.getFilePerNews()))
					'response.write("isNull: " & isNull(objSelNews.getFilePerNews()))
					'response.write("isObject: " & isObject(objSelNews.getFilePerNews()))
					'response.write("isEmpty: " & isEmpty(objSelNews.getFilePerNews()))
					'response.end()
					%>
						
					<%if (not(isNull(objSelNews.getFilePerNews())) AND isObject(objSelNews.getFilePerNews()) AND not(isEmpty(objSelNews.getFilePerNews()))) then
						Set objListaFilePerNews = objSelNews.getFilePerNews()					
						for each xObjFile in objListaFilePerNews
							Set objFileXNews = objListaFilePerNews(xObjFile)
							iTypeFile = objFileXNews.getFileTypeLabel()
							if(Cint(iTypeFile) = 1) then%>	
								<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" hspace="10" vspace="3" border="0" alt="<%=objSelNews.getTitolo()%>" align="top"/>
								<%Exit for
							end if
							Set objFileXNews = nothing	
						next					
						Set objListaFilePerNews = nothing
					end if
										
					Set objSelNews = nothing%>
					</td>
				  </tr>
				<%next%>
			</table>
			<%if(totPages > 1) then%>
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td class="tdContentPaginazione"><p>
					<%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("controller_page"), "")%>
					</p>
				</td>
			</tr>
			</table>
			<%end if%>
		<%else%>
			<br/><br/><div align="center"><b><%=lang.getTranslated("portal.commons.templates.label.page_in_progress")%></b></div>
		<%end if%>
		</td>
		<td class="tdMenuRight">
		<!-- #include virtual="/common/include/MenuContattiDx.inc" -->
		</td>		
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