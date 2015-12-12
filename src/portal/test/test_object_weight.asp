<%
dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject")
set f=fs.CreateTextFile(Server.MapPath("/public")&"\test.txt",true)

Set obj = Session("langElements")

dim a
a=Request.TotalBytes

response.write("total bytes: "&a&"<br><br>")

response.write("obj.count: "&obj.count&"<br><br>")

for each x in obj
	response.write("key: "&x&" -- value: "&obj(x)&"<br>")
	f.Writeline(x&obj(x))
next
f.Close

Set obj =nothing
set f=nothing
set fs=nothing
%>