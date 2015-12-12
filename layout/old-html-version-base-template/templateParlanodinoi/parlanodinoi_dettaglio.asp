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
</head>
<body>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
    <td class="tdContainerTop"><!-- #include file="include/header.inc" --></td>	
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td class="tdContainerContent">	
		<table border="0" cellpadding="0" cellspacing="0" align="left" class="tableContentPdnoi">
		  <tr>
			<td class="parlanoLeft"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" width="194" height="2" vspace="0" hspace="0" border="0"></td>
			<td><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateParlanodinoi/img/sfondo_center.jpg"%>" vspace="0" hspace="0" border="0" name="parlanodinoiRollover" id="parlanodinoiRollover"></td>
			<td class="parlanoRight">			
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
						


						response.write("<b>"&lang.getTranslated("portal.template_parlanodinoi.label.testate")&"</b><br/><br/><br/>")
						

						objTmpNews = objListaNews.Items		
						for newsCounter = FromNews to ToNews
							Set objSelNews = objTmpNews(newsCounter)

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
									Dim hasBigImg, hasPdf, hasAudioVideo, hasOthers
									hasBigImg = false
									hasPdf = false
									hasAudioVideo = false
									hasOthers = false
									
									for each xObjFile in objListaFilePerNews
										Set objFileXNews = objListaFilePerNews(xObjFile)					
										
										select case objFileXNews.getFileTypeLabel()
										case 2
											hasBigImg = true	
										case 3
											hasPdf = true
										case 4
											hasAudioVideo = true
										case 5
											hasOthers = true
										case else
										end select
										Set objFileXNews = nothing	
									next

									if (cbool(hasBigImg)) then
										Dim srcFileBig%>
										<script type="text/javascript" language="JavaScript">
										<!--
										<%' Lista immagini grandi
										for each xObjFile in objListaFilePerNews
											Set objFileXNews = objListaFilePerNews(xObjFile)
											if(objFileXNews.getFileTypeLabel() = 2) then%>	
												Buffer("<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>");
												<%srcFileBig = Application("dir_upload_news")&objFileXNews.getFilePath()
											end if
											Set objFileXNews = nothing											
											'Exit For
										next%>			
										//-->
										</script>
										<a href="#" class="attachLinkStyle" onmouseover="changeBackgroundImg('parlanodinoiRollover','<%=srcFileBig%>');" onmouseout="changeBackgroundImg('parlanodinoiRollover','<%=Application("baseroot")&"/templates/templateParlanodinoi/img/sfondo_center.jpg"%>');"><%response.Write("<span class=titleContent>"&objSelNews.getTitolo()&"</span><br/>")
										response.Write(objSelNews.getTesto()&"<br/>")%></a>
									<%else							
										response.Write("<span class=titleContent>"&objSelNews.getTitolo()&"</span><br/>")
										response.Write(objSelNews.getTesto()&"<br/>")									
									end if
									
									' Lista pdf
									if(hasPdf) then
										for each xObjFile in objListaFilePerNews
											Set objFileXNews = objListaFilePerNews(xObjFile)					
											if(objFileXNews.getFileTypeLabel() = 3) then%>
												<a target="_blank" href="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" class="attachLinkPdfStyle"><%=objFileXNews.getFileDida()%></a><br>					
												<%objListaFilePerNews.remove(xObjFile)
											end if
											Set objFileXNews = nothing	
										next
									end if
								
								end if
								Set objListaFilePerNews = nothing
							else							
								response.Write("<span class=titleContent>"&objSelNews.getTitolo()&"</span><br/>")
								response.Write(objSelNews.getTesto()&"<br/>")
							end if

							response.write("<br/><br/>")


							Set objSelNews = nothing
						next%>				

						<%if(totPages > 1) then%>
							<hr align="center" width="100%">
							<%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("controller_page"), "")
						end if%>
				<%else
					response.Write("<br/><br/><div align=""center""><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
				end if%>
				</td>
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td><!--<img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateParlanodinoi/img/bottom_title_"&lang.getLangCode()&".gif"%>" vspace="0" hspace="0" border="0">--></td>	
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