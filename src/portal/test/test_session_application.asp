<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
</head>
<body>
<h2>Application.Contents</h2>
<%
for each x in Application.Contents
	if(x="GL_Cache_neme-sys")then
		if(isArray(Application.Contents(x)))then
			Set objCache = Application.Contents(x)(0)
			if (strComp(typename(objCache), "Dictionary") = 0)then
				response.write("<br /><b>GL_Cache_neme-sys elements: start</b><br />")
				for each y in objCache
					response.write(y  & "<br />")
				next
				response.write("<b>GL_Cache_neme-sys elements: end</b><br /><br />")
				
			end if
		end if
	else
		Response.Write(x & "=" & Application.Contents(x) & "<br />")
	end if
next%>


<h2>Session.Contents</h2>
<%for each x in Session.Contents
  Response.Write(x & "=" & Session.Contents(x) & "<br />")
next
%> 
</body>
</html>