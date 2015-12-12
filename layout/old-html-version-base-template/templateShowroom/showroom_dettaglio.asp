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
<div id="loading"><img src="<%=Application("baseroot")&"/common/img/loading.gif"%>" vspace="0" hspace="0" border="0" alt="Loading..." width="200" height="50"></div>
<script>
function finish(){
	document.getElementById("loading").style.visibility = "hidden";
	}
</script>
</head>
<body onload="finish()">
</script>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
    <td class="tdContainerTop"><!-- #include virtual="/common/include/header.inc" --></td>	
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td class="tdContainerContent">
		<table border="0" cellpadding="0" cellspacing="0" align="left">
		  <tr>
			<td class="tdContainerLeft">
			<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0" width="337" height="444" id="showroom" hspace="0" vspace="0"> 
			<param name="movie" value="img/flash_showroom.swf"/>  
			<param name="quality" value="high"/> 
			<param name="bgcolor" value="#ffffff"/> 
			<embed src="img/flash_showroom.swf" quality="high" bgcolor="#ffffff" width="337" height="444" name="showroom" align="" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer"></embed> 
			</object></td>
			<td class="tdShowroomContent">			
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
					Set objCurrentNews = News.findNewsByID(id_news)	
				else
					response.Write("<br/><br/><div align=""center""><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
				end if%>
				
				
				<%if(bolHasObj) then%>
					<!-- #include file="include/fileAllegati.inc" -->
					<%Set objCurrentNews = nothing
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
    <td><!--<img src="<%'=Application("baseroot")&Application("dir_upload_templ")&"templateShowroom/img/showroom_bott.gif"%>" vspace="0" hspace="0" border="0">--></td>	
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