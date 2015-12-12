<!-- #include file="IncludeShopObjectList.inc" -->
<%
Dim id_commento, objCommento
id_commento = request("id_commento")

Set objCommento = New CommentsClass
call objCommento.updateStatus(id_commento,1)
Set objCommento = nothing
%>	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
</head>
<body>
	<div id="container">
		<div id="content-popup">	
		<%=lang.getTranslated("frontend.popup.label.comment_posted")%>
		</div>
	</div>
</body>
</html>