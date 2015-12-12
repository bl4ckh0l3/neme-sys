<%
Response.Expires = -1
Server.ScriptTimeout = 1200

On Error Resume Next
Set Upload = Server.CreateObject("Persits.Upload")

' Do not throw the "Wrong ContentType error first time out
'Upload.IgnoreNoPost = True

isObj = typename(Upload)

if(Err.number <> 0)then
	strErr = "<br>Err.number: "&Err.number&"<br>Err.description: "&Err.description
end if
%>

<HTML><title>AspUpload: Test Oggetto Upload</title> 
<BODY BGCOLOR="#FFFFFF">

<h3>AspUpload: Test Upload Object</h3>
<%
response.write("<b>Persits.Upload:</b> "&isObj&"<br>")

response.write("<b>Error:</b> "&strErr&"<br>")
%>
</BODY> 
</HTML>