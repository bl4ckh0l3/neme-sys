<%'On Error Resume Next%>
<!--#include virtual="/common/include/IncludeShopObjectList.inc" -->
<!--include file ="common/include/Objects/objPageCache.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
</head>
<body>
<%
'Dim arrMenuTips, menuTips
'Set menuTips = new MenuClass
'arrMenuTips = menuTips.getTipsMenu(strGerarchia)

	'Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'if not(objFSO.FolderExists("C:\prova")) then
		'response.Write("not exist!")
		'objFSO.CreateFolder("C:\prova")	
	'end if
	
	'if objFSO.FolderExists("C:\prova") then
		'response.Write("exist")
		'objFSO.DeleteFolder "C:\prova"	
	'end if
	'Set objFSO = nothing
	'response.Write("done!")
	
'Dim path
'path = "C:\Inetpub\wwwroot\portal\templates\prova_templates\"
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'objFSO.CreateFolder(path)
'objFSO.DeleteFile(path&"*.*")
'objFSO.DeleteFolder(path&"*.*")
'Set objFSO = nothing

'Dim path
'Set Shell = Createobject("wscript.shell")
'path = "C:\prova\"
'shell.Run "%comspec% /c rd /s/q "&path&"*.*"
'shell.Run "%comspec% /c del /q "&path&"*.*"
'Set Shell = nothing

'Dim sPath
'sPath = "C:\Inetpub\wwwroot\portal\templates\prova_template\"
'Server.CreateObject("WScript.Shell").Run "cmd.exe /c rd """ & sPath & """ /s/q", 0
'Dim objCache
'Set objCache = New CPageCache
'objCache.AutoCacheToMemory()
'Set objCache = Nothing
%>

<% 
'strConnection = "driver={MySQL ODBC 3.51 Driver};server=localhost;uid=user_db;pwd=$fgrtpw_fg6;database=portal;port=3306"
'Set objConn = Server.CreateObject("ADODB.Connection")
'objConn.ConnectionString = strConnection '"DRIVER={MySQL ODBC 3.51 Driver}; UID="&Application("DBUser")&"; PWD="&Application("DBPwd")&"; database="&Application("srt_dbconn")	

'********* SETTO TUTTE LE VARIABILI APPLICATION RECUPERANDOLE DAL DB
'Dim strSQLRs, objRS, strKey, strType, strTmp
'objConn.Open()
'strSQLRs = "SELECT * FROM config_portal"

'Set objRS = objConn.Execute(strSQLRs)		
'f not(objRS.EOF) then
'	do while not objRS.EOF
'		strKey = objRS("keyword")
		'Application(strKey) = objRS("conf_value")
'		response.Write(strKey & "<br>" & objRS("conf_value") & "<br><br>")
'		objRS.moveNext()			
'	loop
'end if
'Set objRS = Nothing
'objConn.Close()	

'response.Write(Application("spese_spedizione_default") & "<br>")
'response.write(Application("srt_dbconn"))

'response.write(Server.MapPath("/mdb-database/db_portal.mdb"))


'response.write("headers: " & Request.ServerVariables("ALL_HTTP"))

'response.write("https: " & Request.ServerVariables("HTTPS"))

%>

<script type="text/javascript">
/*
function verify(xmlDocument) 
{ 
	// 0 Object is not initialized 
	// 1 Loading object is loading data 
	// 2 Loaded object has loaded data 
	// 3 Data from object can be worked with 
	// 4 Object completely initialized 
	if (xmlDocument.readyState != 4) { 
		return false; 
	} 
 } 
 
var text="<bookstore>"
text=text+"<book>";
text=text+"<title>Everyday Italian</title>";
text=text+"<author>Giada De Laurentiis</author>";
text=text+"<year>2005</year>";
text=text+"</book>";
text=text+"</bookstore>";

var xmlDoc, xmlObj;
try //Internet Explorer
  {
  //alert("sono prima di inzializzare!");
  xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
  xmlDoc.async="false";
  //alert("isNull(xmlDoc)"+isNull(xmlDoc));
  //xmlDoc.onreadystatechange=verify(xmlDoc); 
  //alert(xmlDoc.onreadystatechange);
  //xmlDoc.loadXML(text);
  //xmlObj=xmlDoc.documentElement; 
  }
catch(e)
  {
  //try //Firefox, Mozilla, Opera, etc.
  //  {
  //  parser=new DOMParser();
  //  xmlDoc=parser.parseFromString(text,"text/xml");
  //  xmlObj=xmlDoc.documentElement; 
  //  }
  //catch(e) {alert(e.message)}
  //alert("eccezione: " + e);
  }
try 
  {
	document.write("xmlObj "+xmlObj.xml);
  }
catch(e) {//alert(e.message)}*/
</script>
<%
'Dim textXML
'textXML = "<?xml version=""1.0"" encoding=""ISO8859-1"" ?>" &_
'"<bookmarks>" &_
'  "<folder id=""Albums"">" &_
'    "<bookmark href=""http://www.zing.com/album/?id=4294337405"" added=""2000-05-24 10:47:05"">" &_
'      "<title>ZingAlbum: Chewy my temporary dog</title>" &_
'    "</bookmark>" &_
'    "<bookmark href=""http://www.zing.com/album/?id=4294337505"" added=""2000-05-24 10:48:51"">" &_
'      "<title>ZingAlbum: Hanging out @ Victoria Day</title>" &_
'    "</bookmark>" &_
'  "</folder>" &_
'"</bookmarks> "


'Inizializziamo il Parser MS XML...
'Set objXML = Server.CreateObject("Microsoft.XMLDOM")
'objXML.async = False

'Carica il file XML
'strFile = Server.MapPath("bookmarks.xml")
'objXML.LoadXML (textXML)

'Set AllItems = objXML.selectNodes("//folder")

'For I = 0 to (AllItems.Length - 1)
'  Response.Write("<font color="&chr(34)&"#ababab"&chr(34)&">" &_
'  AllItems(I).GetAttribute("id") & "</font><br>" & vbcrlf)
'  Set Bookmarks = AllItems(I).selectNodes("bookmark")
'  Response.Write("<ul>" & vbcrlf)
'  For J = 0 to (Bookmarks.Length-1)
'    Response.Write("<li>" & vbcrlf)
'    Set Title = Bookmarks(J).selectNodes("title")
'    Response.Write("<a href=" & chr(34) & Bookmarks(J).getAttribute("href") & chr(34) & ">")
'    Response.Write(Title(0).text & "</a></li>" & vbcrlf)
'    Set Title = nothing
'  Next
'  Response.Write("</ul>" & vbcrlf)
'  Set Bookmarks = Nothing
'Next

'Dim myDynArray() 'Dynamic size array
'ReDim myDynArray(1)
'response.write("UBound(myDynArray): " & UBound(myDynArray)&"<br>")
'myDynArray(0) = "Albert Einstein"
'myDynArray(1) = "Mother Teresa"
'ReDim Preserve myDynArray(3)
'myDynArray(2) = "Bill Gates"
'myDynArray(3) = "Martin Luther King Jr."
'response.write("UBound(myDynArray): " & UBound(myDynArray)&"<br>")
'For Each item In myDynArray
	'Response.Write(item & "<br />")
'Next

'On Error Resume Next
'Dim strSQLRs, strSQLDelProdCarr, strSQLDel, objRS, strSQLLog, utenteLoggedList, sessionLoggedList, item, objConn
'Dim dta_ins, DD, MM, YY, HH, MIN, SS

'response.write("Session.SessionID: " & Session.SessionID & "<br>")

'strSQLRs = "SELECT id_carrello FROM carrello WHERE id_utente=" & Session.SessionID
'strSQLDelProdCarr = "DELETE FROM prodotti_x_carrello WHERE id_carrello="
'strSQLDel = "DELETE FROM carrello WHERE id_utente=" & Session.SessionID	
'strSQLLog = "INSERT INTO LOGS(msg,usr,type,date_event) VALUES("

'Set objConn = Server.CreateObject("ADODB.Connection")
'objConn.ConnectionString = Application("srt_dbconn")	
'objConn.Open()			
'objConn.BeginTrans	

'Set sessionLoggedList = Server.CreateObject("Scripting.Dictionary")
	
'Set objRS = objConn.Execute(strSQLRs)		
'if not(objRS.EOF) then
'	Dim idCarrello
'	objRS.moveFirst()
'	do while not objRS.EOF
'		idCarrello = objRS("id_carrello")
'		response.write("idCarrello: " & idCarrello & "<br>")
'		sessionLoggedList.add idCarrello, idCarrello
'		objRS.moveNext()
'	loop
'	Set objRS = Nothing
'	response.write(typename(sessionLoggedList)&"<br>")
'	response.write(sessionLoggedList.Count&"<br>")
'	Dim strSQLDelProdCarrTmp
'	For Each item In sessionLoggedList.Keys
'		response.write("item: " & item&"<br>")
'		response.write("strSQLDelProdCarr & item: " & strSQLDelProdCarr & item&"<br>")
'		strSQLDelProdCarrTmp = strSQLDelProdCarr & item
'		objConn.execute(strSQLDelProdCarrTmp)
'	Next
	
'	response.write("strSQLDel: " & strSQLDel&"<br>")
	
'	objConn.execute(strSQLDel)
	
'	response.write("loop done<br>")	

	' registro l'evento nella tabella di logs			
'	dta_ins = Now()
'	DD = DatePart("d", dta_ins)
'	MM = DatePart("m", dta_ins)
'	YY = DatePart("yyyy", dta_ins)
'	HH = DatePart("h", dta_ins)
'	MIN = DatePart("n", dta_ins)
'	SS = DatePart("s", dta_ins)
'	dta_ins = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS

'	strSQLLog = strSQLLog & "'Deleted carrello: " & Session.SessionID & "'"
'	strSQLLog = strSQLLog & ",'system'"
'	strSQLLog = strSQLLog & ",'info'"
'	strSQLLog = strSQLLog & ",'" & dta_ins & "'"
'	strSQLLog = strSQLLog & ")"

'	objConn.execute(strSQLLog)		
'end if
'Set objRS = Nothing
	
'if objConn.Errors.Count = 0 then
'	objConn.CommitTrans
'end If

'if objConn.Errors.Count > 0 then
'	objConn.RollBackTrans
'end if	
						
'objConn.Close()		
'Set sessionLoggedList = nothing
'Set objConn = nothing

'if(Err.number <> 0) then
'	response.write("error= ")
'	response.write(Err.description)
'else
' response.write("esecuzione OK")
'end if

'response.Write(request.ServerVariables("ALL_HTTP"))
response.write("<br/>RETURN NOT OK<br/><br/>")
%> 
</body>
</html>