<%@ LANGUAGE="VBSCRIPT" %>
<!-- #include virtual="/common/include/IncludeShopObjectList.inc" -->

<HTML>
<BODY>
<%
dim cn(10)
dim cmd(10)

For x = 0 to 10
	Set cn(x) = Server.CreateObject("ADODB.Connection")
	cn(x).Open Application("srt_dbconn")
	Set cmd(x) = Server.CreateObject("ADODB.Command")
	cmd(x).activeconnection = cn(x)
	cmd(x).commandtext = "SELECT * FROM news_find"
	
	cmd(x).execute
	
	Response.Write "Command executed: " & x & "<BR>"
	
	
	Set cmd(x) = Nothing
	cn(x).close 'comment this line out to recreate the problem
	Set cn(x) = Nothing
Next

dim cn2(20)
dim cmd2(20)
For x = 0 to 20
	Set cn2(x) = Server.CreateObject("ADODB.Connection")
	cn2(x).Open Application("srt_dbconn")
	Set cmd2(x) = Server.CreateObject("ADODB.Command")
	cmd2(x).activeconnection = cn2(x)
	cmd2(x).commandtext = "SELECT * FROM news_find"
	
	cmd2(x).execute
	
	Response.Write "Command executed: " & x & "<BR>"
	
	
	Set cmd2(x) = Nothing
	cn2(x).close 'comment this line out to recreate the problem
	Set cn2(x) = Nothing
Next

dim cn3(20)
dim cmd3(20)
For x = 0 to 20
	Set cn3(x) = Server.CreateObject("ADODB.Connection")
	cn3(x).Open Application("srt_dbconn")
	Set cmd3(x) = Server.CreateObject("ADODB.Command")
	cmd3(x).activeconnection = cn3(x)
	cmd3(x).commandtext = "SELECT * FROM news_find"
	
	cmd3(x).execute
	
	Response.Write "Command executed: " & x & "<BR>"
	
	
	Set cmd3(x) = Nothing
	cn3(x).close 'comment this line out to recreate the problem
	Set cn3(x) = Nothing
Next

dim cn4(20)
dim cmd4(20)
For x = 0 to 20
	Set cn4(x) = Server.CreateObject("ADODB.Connection")
	cn4(x).Open Application("srt_dbconn")
	Set cmd4(x) = Server.CreateObject("ADODB.Command")
	cmd4(x).activeconnection = cn4(x)
	cmd4(x).commandtext = "SELECT * FROM news_find"
	
	cmd4(x).execute
	
	Response.Write "Command executed: " & x & "<BR>"
	
	
	Set cmd4(x) = Nothing
	cn4(x).close 'comment this line out to recreate the problem
	Set cn4(x) = Nothing
Next

on error Resume Next
Set News = New NewsClass
for counter = 1 to 50
	Set objListaNews = News.findNews(null, null, null, null, null, null, null, null, null, null)
	response.write("objListaNews: "&counter&" - "&typename(objListaNews)&"<br>")
	Set objListaNews = nothing
next
Set News = nothing


'Set Order = New OrderClass
'for counter2 = 1 to 50
'	Set objListaOrder = Order.findOrdini(null, null, null, null, null, null, null, null, null)
'	response.write("objListaOrder: "&counter2&" - "&typename(objListaOrder)&"<br>")
'	Set objListaOrder = nothing
'next
'Set Order = nothing

if Err.number <> 0 then
	response.write(Err.description)
end if	
%>
</BODY>
</HTML>