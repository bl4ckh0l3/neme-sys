<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->
<!-- #include file="include/init2.inc" -->
<!-- #include virtual="/common/include/setTemplateTargetList.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="include/initMeta2.inc" -->
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<!-- #include file="include/initStyleAndJs2.inc" -->
</head>
<body onload="initialize()" onunload="GUnload()">
<div id="warp">
	<!-- #include virtual="/public/layout/include/header.inc" -->	
	<div id="container">	
		<!-- include virtual="/public/layout/include/menu_orizz.inc" -->
<!-- #include virtual="/public/layout/include/menu_vert_sx.inc" -->
		<div id="content-center">
			<!-- #include virtual="/public/layout/include/menutips.inc" -->
			<!-- #include file="include/initContent2.inc" -->
		</div>
		<!-- #include virtual="/public/layout/include/menu_vert_dx.inc" -->
		<!-- #include virtual="/public/layout/addson/contents/news_comments_widget.inc" -->
		<!-- Il file attachments.inc si aspetta sempre valorizzato: id_prodotto -->
		<%if(bolHasObj) then%>
			<!-- #include file="include/attachments.inc" -->
			<%Set objCurrentNews = nothing
		end if%>
	</div>
	<!-- #include virtual="/public/layout/include/bottom.inc" -->
</div>
</body>
</html>
<!-- #include file="include/end2.inc" -->