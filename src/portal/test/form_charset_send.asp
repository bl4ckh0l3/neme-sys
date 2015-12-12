<%'@Language=VBScript codepage=65001 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=utf-8">-->
<%'Response.Charset="UTF-8"%>
</head>
<body>
<form action="form_charset_receive.asp" method="post" name="form_inserisci"><!--  accept-charset="UTF-8"  enctype="multipart/form-data" -->
<input type="text" name="codice_prod" value=""><!--  accept="text/plain;charset=UTF-8" -->
<input type="submit" value="invia" name="send"/>			
</form>
</body>
</html>