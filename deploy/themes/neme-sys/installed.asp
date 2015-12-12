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
<!-- inizio container -->
<div id="container">
	<!-- header -->
	<div id="header" class="bg_nemesys">
		<div class="header_navbar">
			<h1 class="nemesys_h1" title="nemesys cms">ne<span>me-sys</span></h1>			
			<div class="div_login">
				<a href="<%=Application("baseroot")&"/login.asp"%>">Accedi!</a> Oppure <a href="<%=Application("baseroot")&"/area_user/manageUser.asp"%>">Registrati</a>
			</div>
		</div>		
	</div>
	<!-- header fine -->
	<!-- main -->
	<div id="main">		
		<!-- content -->	
		<div class="content">		
			<form name="install_login" method="post" action ="<%=Application("baseroot")&"/common/include/verificautente.asp"%>">
			<p align="center">
			Il database e' stato correttamente istallato!<br/><br/>
			All query have been executed!<br/><br/>
			
			<input type="hidden" name="j_username" value="administrator">
			<input type="hidden" name="j_password" value="admin">
			<input type="submit" value="LOGIN BACK OFFICE" align="center">
			</p>
			</form>
		
		</div>
		<!-- content fine -->		
	</div>
	<!-- main fine -->	
</div>
<!-- fine container -->
<div id="footer"><span>Powered by BHN Online Technology Merchant Copyright &copy; 2007-2012</span></div>
</body>
</html>