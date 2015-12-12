<%@LANGUAGE="VBSCRIPT"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Documento senza titolo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<HTML>
<BODY>
<%
Response.Status="301 Moved Permanently"
'Response.AddHeader "Location", "http://www.blackholenet.com/test/test_301/test_301_end.asp"
Response.AddHeader "Location", "/test/test_301/test_301_end.asp"
%>
</BODY>
</HTML>
</body>
</html>
