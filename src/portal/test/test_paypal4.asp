<%
Response.ContentType = "text/xml"
if(request.form("id_ordine")=84) then
	response.write("<Result>true</Result>")
else
	response.write("<Result>false</Result>")
end if
%>