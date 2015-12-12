<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>test page</title>
</head>
<body>
<div id="warp">
<%
publicDirVar = "/public"
nemesiConfigFile = publicDirVar & "/conf/test_config.xml"

On Error Resume Next	
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set configFile=objFSO.OpenTextFile(Server.MapPath(nemesiConfigFile), 2, True)
configFile.writeLine("<config>")		
configFile.writeLine("<tag1></tag1>")	
configFile.writeLine("<tag2></tag2>")	
configFile.writeLine("</config>")	
configFile.Close
Set configFile=Nothing
Set objFSO = nothing

If Err.Number<>0 then
	response.Write(Err.description&"<br/><br/>")
end if
%>
</div>
</body>
</html>
