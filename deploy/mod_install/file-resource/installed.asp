<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>installed page</title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
</head>
<body>
<div id="warp">

	<div id="header">
		<div id="top-bar">
			<div id="top-bar-logo"></div>
			<div id="top-bar-search"></div>
			<div id="top-bar-lenguage">
				<ul>
				<li></li>
				</ul>
			</div>
		</div>
		<div id="image-container"></div>
	</div>
	<div id="container">    	
		<div id="menu-left"></div>
		<div id="content-center">
		
		<form name="install_login" method="post" action ="<%=Application("baseroot")&"/common/include/verificautente.asp"%>">
		<p align="center">
		Il database e' stato correttamente istallato!<br/><br/>
		
		<input type="hidden" name="j_username" value="administrator">
		<input type="hidden" name="j_password" value="admin">
		<input type="submit" value="LOGIN" align="center">
		</p>
		</form>
		
		</div>
		<div id="menu-right"></div>
	</div>
	<div id="footer"><span>Powered by BHN Online Technology Merchant Copyright &copy; 2007-2012</span></div>
</div>
</body>
</html>