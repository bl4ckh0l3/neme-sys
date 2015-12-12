<%'On Error Resume Next%>

<% Response.Buffer = true %> 

<!--include virtual="/common/include/IncludeShopObjectList.inc" -->
<!--include file ="common/include/Objects/objPageCache.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%'=lang.getTranslated("frontend.page.title")%></title>
</head>
<body>
<%
	response.Write("ALL_HTTP: "&request.ServerVariables("ALL_HTTP")&"<br>")
	response.Write("REMOTE_ADDR: "&request.ServerVariables("REMOTE_ADDR")&"<br>")
	response.Write("REMOTE_HOST: "&request.ServerVariables("REMOTE_HOST")&"<br>")
	response.Write("URL: "&request.ServerVariables("URL")&"<br>")
	response.Write("HTTP_HOST: "&request.ServerVariables("HTTP_HOST")&"<br>")
	response.Write("SERVER_NAME: "&request.ServerVariables("SERVER_NAME")&"<br>")
	response.Write("SERVER_PORT: "&request.ServerVariables("SERVER_PORT")&"<br>")
	response.Write("SERVER_PROTOCOL: "&request.ServerVariables("SERVER_PROTOCOL")&"<br>")
	response.Write("SCRIPT_NAME: "&request.ServerVariables("SCRIPT_NAME")&"<br>")
	response.Write("PATH_INFO: "&request.ServerVariables("PATH_INFO")&"<br>")
	response.Write("URL: "&request.ServerVariables("URL")&"<br><br><br>")
	
	
for each x in Request.ServerVariables
  response.write(x & ": "&Request.ServerVariables(x)&"<br />")
next	

%> 



<% 
    url = "www.google.com" 
 
    Set objWShell = CreateObject("WScript.Shell") 
    Set objCmd = objWShell.Exec("ping " & url) 
    strPResult = objCmd.StdOut.Readall() 
    set objCmd = nothing: Set objWShell = nothing 
 
    strStatus = "offline" 
    if InStr(strPResult,"TTL=")>0 then strStatus = "online" 
 
    response.write url & " is " & strStatus 
    response.write ".<br>" & replace(strPResult,vbCrLf,"<br>") 
%>
</body>
</html>