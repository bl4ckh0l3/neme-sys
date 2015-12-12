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

Dim objCategoriaTmp, categoriaTmpDesc
Set objCategoriaTmp = new CategoriaClass
categoriaTmpDesc = objCategoriaTmp.findCategoriaByGerarchia(strGerarchia).getCatDescrizione()
Set objCategoriaTmp = nothing
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
<SCRIPT>
<!--
/**********************************************   PRIMA FUNZIONE SLIDESHOW   **********************************************/


var pictures = new Array();
var chooice = "";

var defaultImg = new Image();
defaultImg.id="SlideShow";
defaultImg.name="SlideShow";
defaultImg.src="<%=Application("baseroot")&"/common/img/spacer.gif"%>";


<%	
Dim bolHasObj, inProgress
bolHasObj = false
inProgress = ""

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
	<%if not(isNull(objCurrentNews.getFilePerNews())) AND not(isEmpty(objCurrentNews.getFilePerNews())) then
		Set objListaFilePerNews = objCurrentNews.getFilePerNews()
		
		if not(isEmpty(objListaFilePerNews)) then
			' LEGENDA TIPI FILE
			'1 = img small
			'2 = img big
			'3 = pdf
			'4 = audio-video
			'5 = others			
			
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
			
			  Dim counter
						if (cbool(hasBigImg)) then
			  ' Lista immagini grandi
			  counter = 0%>  				
			    <%for each xObjFile in objListaFilePerNews
			      Set objFileXNews = objListaFilePerNews(xObjFile)					
			      if(objFileXNews.getFileTypeLabel() = 2) then%>
		 
					pictures[<%=counter%>] = '<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>';           
					<%if(counter = 0) then%>
					defaultImg.src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>";
					<%end if
				objListaFilePerNews.remove(xObjFile)
			      end if
			      Set objFileXNews = nothing	
			      counter = counter + 1
			    next%>        
			      

			      var preLoad = new Array()
			      for (i = 0; i < pictures.length; i++){
				 preLoad[i] = new Image()
				 preLoad[i].src = pictures[i];              
			      }   
			     
			    var arrSlideshowFunction = new Array(1);
			    arrSlideshowFunction[0]="progid:DXImageTransform.Microsoft.Fade(Overlap=0.3,duration=1,enabled=true)";
			    //arrSlideshowFunction[1]="progid:DXImageTransform.Microsoft.Iris(Overlap=0.1,irisstyle=CROSS,motion=in,duration=1,enabled=true)";
			    //arrSlideshowFunction[2]="progid:DXImageTransform.Microsoft.RadialWipe(wipestyle=CLOCK,duration=2,enabled=false)";
			    //arrSlideshowFunction[3]="progid:DXImageTransform.Microsoft.Wheel(spokes=4,duration=2,enabled=false)";
			    //arrSlideshowFunction[4]="progid:DXImageTransform.Microsoft.Zigzag(GridSizeX=8,GridSizeY=8,duration=2,enabled=false)";
			    //arrSlideshowFunction[5]="progid:DXImageTransform.Microsoft.Pixelate(MaxSquare=50,duration=2,enabled=false)";
			    //arrSlideshowFunction[6]="progid:DXImageTransform.Microsoft.Spiral(GridSizeX=8,GridSizeY=8,duration=2,enabled=false)";
			    
			    //var functChooiceNum = Math.floor(Math.random()*2+0);
			    var functChooiceNum = 0;
			    chooice = arrSlideshowFunction[functChooiceNum];       

			    //alert("functChooiceNum: " + functChooiceNum+ " chooice: " + chooice);	         
			<%end if

		end if
		Set objListaFilePerNews = nothing
	end if%>
	<%Set objCurrentNews = nothing			
else
	'response.Write("<div align=""center""><br/><br/><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>")
	response.write("")
	inProgress = "<div align=""center""><br/><br/><b>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</b></div>"
end if%>


/**********************************************   INIZIO FUNZIONE SLIDESHOW   **********************************************/

a = 0;

function ejs_img_fx(img){
	if(img && img.filters && img.filters[0]){
		img.filters[0].apply();
		img.filters[0].play();
	}
}

function StartAnim(filterType){
	if (document.images){
		var imgElement = document.getElementById("SlideShow");
		imgElement.style.filter=filterType;
		defilimg()
	}
}
	
function defilimg(){
	if (a == pictures.length){
		a = 0;
	}
	
	if (document.images){
		var imgElement = document.getElementById("SlideShow");
		ejs_img_fx(imgElement)
		imgElement.src = pictures[a];
		tempo3 = setTimeout("defilimg()",3000);
		a++;
	}
}

/**********************************************   FINE FUNZIONE SLIDESHOW   **********************************************/

function callSlideShowFunct(){
	if(chooice != "") eval("StartAnim('"+chooice+"');"); 
}
//-->
</SCRIPT>
<div id="loading2"><img src="<%=Application("baseroot")&"/common/img/loading.gif"%>" vspace="0" hspace="0" border="0" alt="Loading..." width="200" height="50"></div>
<script>
function finish(){
	document.getElementById("loading2").style.visibility = "hidden";
	}
</script>
</head>
<body onload="javascript:callSlideShowFunct();finish();">
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
    <td class="tdContainerTop"><!-- #include file="include/header.inc" --></td>	
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td class="tdContainerContent"><SCRIPT>document.write("<img src="+defaultImg.src+" align=\"center\" width=\"955\" height=\"389\" hspace=\"0\" vspace=\"0\" border=\"0\" name=\"SlideShow\" id=\"SlideShow\">");</SCRIPT>
    <%if not(inProgress = "") then response.write(inProgress) end if%></td>
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td><!--<img src="<%'=Application("baseroot")&"/common/img/"&lang.getTranslated("frontend.menu.label.img.bott."&categoriaTmpDesc)&".gif"%>" vspace="0" hspace="0" border="0">--></td>	
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