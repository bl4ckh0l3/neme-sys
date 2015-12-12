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
    <td class="tdContainerTop"><!-- #include file="include/header.inc" --></td>	
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td class="tdContainerContentIndex">			
	<table border="0" cellpadding="0" cellspacing="0" align="left">
	<tr>
	<td align="left" valign="top"><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/listenozze_center.jpg"%>" align="top" vspace="0" hspace="0" border="0"></td>
	<td align="left" valign="top" class="tdContainerContentIndexList">
	<table border="0" cellpadding="0" cellspacing="0" align="left">
	<%
	Set menuFruizione = new MenuFruizioneClass
	Set menuComplete = menuFruizione.getCompleteMenuByMenu("2")
	Set categoriaClassTmp = new CategoriaClass	
	Dim desc_bundle
	
	if(isNull(strGerarchia) OR strGerarchia = "") then strGerarchia = "00"

	iGerLevel = menuFruizione.getLivello(strGerarchia)
	
	for each x in menuComplete
		level = menuFruizione.getLivello(x)
		iGerDiff = level - iGerLevel
	  	menuCompleteCatDescTrans = "frontend.menu.label."&menuComplete(x)
					
		if not(isNull(lang.getTranslated(menuCompleteCatDescTrans))) AND not(lang.getTranslated(menuCompleteCatDescTrans) = "") then
			desc_bundle = lang.getTranslated(menuCompleteCatDescTrans)
		else 
			desc_bundle = menuComplete(x)
		end if
			
		if(level > 1) then
			iWidth = (level-1) * 10 
			iLenGer = (level * 2) + (level -1)
			strSubTmpGer = Left(strGerarchia, iLenGer)
			strSubTmpGerFiltered = Left(strSubTmpGer, iLenGer-3)
			
			if(iGerDiff <= 1) then
				if (InStr(1, Left(x, iLenGer-3), strSubTmpGerFiltered, 1) > 0) then
					hrefGer = x
					strHref = Application("baseroot") & "/common/include/Controller.asp?gerarchia="&hrefGer				
					
					'*** Controllo se la categoria contiene news, altrimenti cerco la prima sottocategoria che contenga news
					'*** e imposto la nuova gerarchia come parametro nel link
					Set objCategoriaCheck = categoriaClassTmp.findCategoriaByGerarchia(hrefGer)
					if not(objCategoriaCheck.contieneNews()) AND not(objCategoriaCheck.contieneProd()) then
						foundGer = categoriaClassTmp.findFirstSubCategoriaWithNews(hrefGer)
						
						if(isNull(foundGer)) then
							foundGer = categoriaClassTmp.findFirstSubCategoriaWithProd(hrefGer)
						end if
						
						if not(isNull(foundGer)) then
							strHref = Application("baseroot") & "/common/include/Controller.asp?gerarchia="&foundGer
						else
							strHref = "#"
						end if
					end if
					Set objCategoriaCheck = nothing%>		
					<tr>
					<td><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/menu_arrow_sx.gif"%>" align="top" vspace="0" hspace="0" border="0"></td>
					<td class="tdMenuListCenter"><a class="linkListeNozzeList" href="javascript:openWin('<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&hrefGer%>','listenozze',screenW,screenH,0,0)"><%=desc_bundle%></a></td>
					<td><img src="<%=Application("baseroot")&Application("dir_upload_templ")&"templateListenozze/img/menu_bg_dx.gif"%>" align="top" vspace="0" hspace="0" border="0"></td>
					</tr>
					<tr><td colspan="3" height="37"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="37" vspace="0" hspace="0" border="0" align="top" ></td></tr>	
			<%end if
			end if
		end if	
	next
	
	Set categoriaClassTmp = nothing
	Set menuComplete = nothing
	Set menuFruizione = nothing
	%>
	</table>
	</td>
	</tr>
	</table>
	</td>
  </tr>
  <tr>
    <td class="trWhite"><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="2" vspace="0" hspace="0" border="0"></td>	
  </tr>
  <tr>
    <td class="tdBottom">&nbsp;</td>	
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