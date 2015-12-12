<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<html>
<head>
</head>
<body>
<%
On Error Resume Next
Dim objCacheClass
Set objCacheClass = new CacheClass
'objCacheClass.clear
call objCacheClass.removeByPrefix("findc", 25)
'call objCacheClass.removeByPrefix("findc", null)

'Class AdsClass
'	public test_prop
'End Class

Set objNews = new NewsClass
Set guidObj = new GUIDClass

'Set testObj = new AdsClass
'testObj.test_prop="prova properties"
'response.write("testObj.test_prop: "&testObj.test_prop&"<br>")

'test_arr = array(server.createObject("scripting.dictionary"))


'Application("obj_test") = test_arr
'Application("obj_test")(0).add "test", testObj

'response.write("Application(obj_test): "&Application("obj_test").test_prop)
'response.write("typename(Application(obj_test)): "&typename(Application("obj_test"))&"<br>")
'response.write("typename: "&typename(Application("obj_test")(0).item("test"))&"<br>")
'response.write("value: "&Application("obj_test")(0).item("test").test_prop)


'Set Session("obj_test") = Server.CreateObject("Scripting.Dictionary")
'Session("obj_test").add "test", testObj
'response.write("typename(Session(obj_test)): "&typename(Session("obj_test"))&"<br>")
'response.write("value: "&Session("obj_test").item("test").test_prop)

'caching.removeAll()

'caching.add guidObj.CreateOrderGUIDLong(), objNews.findNewsByID("25")
'caching.add guidObj.CreateOrderGUIDLong(), objNews.findNewsByID("65")
'caching.add guidObj.CreateOrderGUIDLong()&"", objNews.findNewsByID("67")




'for each q in caching
	'response.write("key: "&q&" - value: "&caching(q).getTitolo()&"<br>")
'next

'caching.removeAll()

if(Err.number<>0)then
response.write(Err.description)
end if
%> 
</body>
</html>
