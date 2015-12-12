<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/area_user.css"%>" type="text/css">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
</head>
<body>
<!-- #include file="grid_top.asp" -->

		<h1><%=lang.getTranslated("frontend.registration.manage.label.confirm_registration")%></h1>
		<%if(Application("confirm_registration") = 2) then%>
		<h1><%=lang.getTranslated("frontend.registration.manage.label.confirm_registration_with_code")%></h1>
		<%end if%>	
		   
<!-- #include file="grid_bottom.asp" -->
</body>
</html>
