<title><%
if not(pageTemplateTitle = "") then 
	if not(lang.getTranslated(pageTemplateTitle) = "") then
		response.write(lang.getTranslated(pageTemplateTitle)) 
	else
		response.write(pageTemplateTitle) 
	end if
else 
	response.Write(lang.getTranslated("frontend.page.title"))
end if
%></title>

<META name="description" CONTENT="<% 
if not(lang.getTranslated(metaDescription) = "") then
	response.write(lang.getTranslated(metaDescription)) 
else
	response.write(metaDescription) 
end if
%>">
<META name="keywords" CONTENT="<% 
if not(lang.getTranslated(metaKeyword) = "") then
	response.write(lang.getTranslated(metaKeyword)) 
else
	response.write(metaKeyword) 
end if
%>">

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">