<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
Dim id_element, element_type, active, objCommento, objSelectedCommento
id_element = request("id_element")
element_type = request("element_type")
active = request("active")
Set objCommento = New CommentsClass

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
</head>
<body>
	<div id="container">	
		<div id="content-popup">
		<%
		if(not(isNull(objCommento.findCommentiByIDElement(id_element,element_type,active)))) AND (Instr(1, typename(objCommento.findCommentiByIDElement(id_element,element_type,active)), "dictionary", 1) > 0) then
			Dim  x, objTmpCommento		
			Set objSelectedCommento = objCommento.findCommentiByIDElement(id_element,element_type,active)
			
			for each x in objSelectedCommento.Keys
				Set objTmpCommento = objSelectedCommento(x)
				response.write(objTmpCommento.getDtaInserimento()&"<br>")
				response.write(objTmpCommento.getMessage()&"<br><hr>")
			next
		else
			response.Write("<div align='center'>"&lang.getTranslated("frontend.popup.label.no_comments")&"</div><br>")
		end if%>
		<div align="center" style="padding-top:30px;"><a href="javascript:window.close();"><%=lang.getTranslated("frontend.popup.label.close_window")%></a></div><br>
		</div>
	</div>
</body>
</html>
<%
Set objSelectedCommento = nothing
Set objCommento = nothing%>