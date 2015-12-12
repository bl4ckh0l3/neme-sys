<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->
<!-- #include file="include/init1.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=pageTemplateTitle%></title>
<META name="description" CONTENT="<%=metaDescription%>">
<META name="keywords" CONTENT="<%=metaKeyword%>">
<META name="autore" CONTENT="Neme-sys; email:info@neme-sys.org">
<META http-equiv="Content-Type" CONTENT="text/html; charset=utf-8">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<%if not(isNull(strCSS)) ANd not(strCSS = "") then%><link rel="stylesheet" href="<%=Application("baseroot") & strCSS%>" type="text/css"><%end if%>
</head>
<body>
<div id="warp">
	<!-- #include virtual="/public/layout/include/header.inc" -->	
	<div id="container">	
		<!-- include virtual="/public/layout/include/menu_orizz.inc" -->
		<!-- #include virtual="/public/layout/include/menu_vert_sx.inc" -->
		<div id="content-center">
			<!-- #include virtual="/public/layout/include/menutips.inc" -->

			<div align="left">
			<%if (bolHasObj) then%>
				<div>
				<p><strong><%=objCurrentNews.getTitolo()%></strong></p>
				<%
				if (Len(objCurrentNews.getAbstract1()) > 0) then response.Write(objCurrentNews.getAbstract1()) end if
				if (Len(objCurrentNews.getAbstract2()) > 0) then response.Write(objCurrentNews.getAbstract2()) end if
				if (Len(objCurrentNews.getAbstract3()) > 0) then response.Write(objCurrentNews.getAbstract3()) end if
				response.Write(objCurrentNews.getTesto())
				
				if(bolHasAttach) then 
					for each key in attachMap
						if(attachMap(key).count > 0)then%>
							<br/><br/><strong><%=lang.getTranslated(attachMultiLangKey(key))%></strong><br/>
							<%for each item in attachMap(key)%>
								<a href="javascript:openWin('<%=Application("baseroot")&"/public/layout/include/popup.asp?id_allegato="&item.getFileID()&"&parent_type=1"%>','popupallegati',400,400,100,100)"><%=item.getFileName()%></a><br>
							<%next
						end if
					next
				end if
				Set objCurrentNews = nothing				
				%>
				</div>
				<div id="torna"><a href="<%=Application("baseroot") & "/common/include/feedRSS.asp?gerarchia="&strGerarchia&"&id_news="&id_news&"&page="&numPage&"&modelPageNum="&modelPageNum%>" target="_blank"><img src="<%=Application("baseroot")&"/common/img/rss_image.gif"%>" vspace="3" hspace="3" border="0" align="right" alt="RSS"></a></div>
			<%else%>
				<br/><br/><div align="center"><strong><%=lang.getTranslated("portal.commons.templates.label.page_in_progress")%></strong></div>
			<%end if%>
			</div>
			
			
			<!-- commenti -->
			<div id="comments-widgets">
				<div id="fb-root"></div>
				<script>(function(d, s, id) {
				var js, fjs = d.getElementsByTagName(s)[0];
				if (d.getElementById(id)) {return;}
				js = d.createElement(s); js.id = id;
				js.src = "//connect.facebook.net/en_US/all.js#xfbml=1";
				fjs.parentNode.insertBefore(js, fjs);
				}(document, 'script', 'facebook-jssdk'));</script>

				<div class="fb-comments" data-href="domain.com" data-num-posts="3" data-width="500"></div>

				<br/><br/><br/>
				<script src="//platform.twitter.com/widgets.js" type="text/javascript"></script>
				<div><a href="https://twitter.com/share" class="twitter-share-button" data-count="horizontal">Tweet</a></div>       
			</div>       
			<!-- commenti fine -->			
			
		</div>
		<!-- #include virtual="/public/layout/include/menu_vert_dx.inc" -->
		<!-- #include virtual="/public/layout/addson/contents/news_comments_widget.inc" -->
	</div>
	<!-- #include virtual="/public/layout/include/bottom.inc" -->
</div>
</body>
</html>
<%
'****************************** PULIZIA DEGLI OGGETTI UTILIZZATI
Set objListaTargetCat = nothing
Set objListaTargetLang = nothing
Set objListaNews = nothing
Set News = Nothing
%>