<%On Error Resume Next%>
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


'response.Write(request.ServerVariables("ALL_HTTP"))
'Response.AddHeader("REQUEST_METHOD","POST")
'response.Write(request.ServerVariables("REQUEST_METHOD"))
%> 




<!--<form action="https://api.sandbox.paypal.com/nvp/" method="post">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="denis.testa@gmail.com">
<input type="hidden" name="item_name" value="carrello:dgtshtgsfa3454whgbg">
<input type="hidden" name="currency_code" value="EUR">
<input type="hidden" name="amount" value="0.01">
<input type="image" src="http://www.paypal.com/it_IT/i/btn/x-click-but01.gif" name="submit" alt="Effettua i tuoi pagamenti con PayPal. È un sistema rapido, gratuito e sicuro.">
</form>-->

<!--<form action="https://www.sandbox.paypal.com/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_s-xclick">
<input type="hidden" name="hosted_button_id" value="28848">
<input type="image" src="https://www.sandbox.paypal.com/en_US/i/btn/btn_buynowCC_LG.gif" border="0" name="submit" alt="PayPal - The safer, easier way to pay online!">
<img alt="" border="0" src="https://www.sandbox.paypal.com/en_US/i/scr/pixel.gif" width="1" height="1">
</form>-->


<!--
<form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="_blank">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="denis.testa@gmail.com">
<input type="hidden" name="rm" value="2">
<input type="hidden" name="return" value="http://www.blackholenet.com/test/test_paypal_return.asp">
<input type="hidden" name="cancel_return" value="http://www.blackholenet.com/test/test_paypal_cancel_return.asp">
<input type="hidden" name="item_name" value="carrello:dgtshtgsfa3454whgbg">
<input type="hidden" name="amount" value="0.01">
<input type="hidden" name="currency_code" value="EUR">
<input type="hidden" name="custom" value="codicecontrollo:4564364335345">
<input type="hidden" name="invoice" value="ordine:46682944">
<input type="hidden" name="image_url" value="http://www.blackholenet.com/common/img/dsm/logo_dsm.gif">
<input type="hidden" name="ack" value="">
<input type="image" src="http://www.paypal.com/it_IT/i/btn/x-click-but01.gif" name="submit" alt="Effettua i tuoi pagamenti con PayPal. È un sistema rapido, gratuito e sicuro.">
</form>


<form method="post" action="https://api-3t.sandbox.paypal.com/nvp">
<input type=hidden name=USER value="denis._1237587782_biz_api1.gmail.com">
<input type=hidden name=PWD value="1237587792">
<input type=hidden name=SIGNATURE value="ArFz1VY1eQ1xhXkf.lQp952AtkMxAt5zep55SUdhzKhEdLC.RdVC0saL">
<input type=hidden name=VERSION value="2.3">
<input type=hidden name=PAYMENTACTION value="Authorization">
<input name=AMT value="29.95">
<input type=hidden name=RETURNURL value="http://www.blackholenet.com/test/test_paypal_return.asp">
<input type=hidden name=CANCELURL value="http://www.blackholenet.com/test/test_paypal_cancel_return.asp">
<input type=submit name=METHOD value="SetExpressCheckout">
</form>

<form method=post action="https://api-3t.sandbox.paypal.com/nvp">
<input type=hidden name=USER value="denis._1237587782_biz_api1.gmail.com">
<input type=hidden name=PWD value="1237587792">
<input type=hidden name=SIGNATURE value="ArFz1VY1eQ1xhXkf.lQp952AtkMxAt5zep55SUdhzKhEdLC.RdVC0saL">
<input type=hidden name=VERSION value="2.3">
<input name="TOKEN" value="EC-4DK82969PW647291T">
<input type=submit name=METHOD value="GetExpressCheckoutDetails">
</form>

<form method=post action="https://api-3t.sandbox.paypal.com/nvp">
<input type=hidden name=USER value="denis._1237587782_biz_api1.gmail.com">
<input type=hidden name=PWD value="1237587792">
<input type=hidden name=SIGNATURE value="ArFz1VY1eQ1xhXkf.lQp952AtkMxAt5zep55SUdhzKhEdLC.RdVC0saL">
<input type=hidden name=VERSION value="2.3">
<input type=hidden name=PAYMENTACTION value="Authorization">
<input type=hidden name=PAYERID value="TPBHRGFMNTXZY">
<input type=hidden name=TOKEN value="EC-4DK82969PW647291T">
<input type=hidden name=AMT value="29.95">
<input type=submit name=METHOD value="DoExpressCheckoutPayment">
</form>

-->

<form method=post action="http://localhost/common/include/checkin.asp">
<input type=hidden name=ack value="Success">
<input type=hidden name=custom value="0000002911680000562743178CGN5EFOLRO5NB3ZW026MOHPYCXC85W194ZD2JKHIXDAIZHJB2RBPY7ZRYZF93XW6TMHM1E1N0M6MEW8AYN0EM0AT3RJIP6WVKTF4FN4ZDBPO5S42D7V6BUUVPJCSX|85|16.00">
<input type=hidden name=amount value="16.00">
<input type=submit name=METHOD value="pay">
</form>
<%
'response.write("<br/><br/>Variabile in sessione Pippo: "& Session("pippo")&"<br/>")
'response.write("<br/>Variabile in sessione amount: "& Session("amount")&"<br/><br/>")

' read post from PayPal system and add 'cmd'
'str = Request.Form & "&cmd=_notify-validate"
'str = "amount=10"

' post back to PayPal system to validate
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP30")
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP40")
'response.Write("typename(objHttp): "&typename(objHttp)&"<br/><br/>")
'objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
'objHttp.open "POST", "http://localhost/test/test_paypal2.asp", false
'objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
'objHttp.Send(str)
'objHttp.Send()

'response.Write("objHttp.getAllResponseHeaders(): "&objHttp.getAllResponseHeaders()&"<br/><br/>")
'response.Write("objHttp.responseText(): "&objHttp.responseText()&"<br/><br/>")
'response.Write("objHttp.responseXML(): "&objHttp.responseXML()&"<br/><br/>")

'response.Write("objHttp.responseBody: "&objHttp.responseBody()&"<br/><br/>")
'response.write("<br/><br/>Variabile in sessione Pippo: "& Session("pippo")&"<br/><br/>")
'response.write("<br/><br/>Variabile in sessione amount: "& Session("amount")&"<br/><br/>")

' assign posted variables to local variables
'Item_name = Request.Form("item_name")
'Item_number = Request.Form("item_number")
'Payment_status = Request.Form("payment_status")
'Payment_amount = Request.Form("mc_gross")
'Payment_currency = Request.Form("mc_currency")
'Txn_id = Request.Form("txn_id")
'Receiver_email = Request.Form("receiver_email")
'Payer_email = Request.Form("payer_email")

' Check notification validation
'if (objHttp.status <> 200 ) then
' HTTP error handling
'elseif (objHttp.responseText = "VERIFIED") then
' check that Payment_status=Completed
' check that Txn_id has not been previously processed
' check that Receiver_email is your Primary PayPal email
' check that Payment_amount/Payment_currency are correct
' process payment
'elseif (objHttp.responseText = "INVALID") then
' log for manual investigation
'else
' error
'end if
'set objHttp = nothing


'************************    START TEST CHIAMATA A LOCALE E VISUALIZZAZIONE XML O TESTO RISPOSTA   ********************
'str = "amount=10"
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'objHttp.open "POST", "http://localhost/test/test_paypal2.asp", false
'objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
'objHttp.Send(str)
'response.write(objHttp.statusText&"<br>")
'response.write(objHttp.status&"<br>")
'response.Write(objHttp.responseXML.xml&"<br><br/>")
'response.Write(objHttp.responseText&"<br/><br/>")
'set objHttp = nothing

'Session("payment_opt") = "amount=10&id_ordine=12365&checkid=3452pjfalfj4l5j43qrh"
'pageRedirect = Application("baseroot") & "/test/test_paypal3.asp"
''Server.Transfer(pageRedirect)
'response.Redirect(pageRedirect)
'************************    FINISH TEST CHIAMATA A LOCALE E VISUALIZZAZIONE XML O TESTO RISPOSTA   ********************


'************************    START TEST CHIAMATA A LOCALE E VISUALIZZAZIONE XML O TESTO RISPOSTA 2   ********************
'str = "id_ordine=84"
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'objHttp.open "POST", "http://localhost/test/test_paypal4.asp", false
'objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
'objHttp.Send(str)
'Set objXML = objHTTP.ResponseXML
'set items = objXML.getElementsByTagName("Result")
'val = items(0).childNodes(0).nodeValue
'response.write("boolean: "&Cbool(val)&"<br>")
'set items = nothing
'Set objXML = nothing
'set objHttp = nothing

'Session("payment_opt") = "amount=10&id_ordine=12365&checkid=3452pjfalfj4l5j43qrh"
'pageRedirect = Application("baseroot") & "/test/test_paypal3.asp"
''Server.Transfer(pageRedirect)
'response.Redirect(pageRedirect)
'************************    FINISH TEST CHIAMATA A LOCALE E VISUALIZZAZIONE XML O TESTO RISPOSTA 2   ********************


'************************    START TEST CHIAMATA A PAYPAL SetExpressCheckout E VISUALIZZAZIONE RISPOSTA   ********************
'str ="USER=denis._1237587782_biz_api1.gmail.com&PWD=1237587792&SIGNATURE=ArFz1VY1eQ1xhXkf.lQp952AtkMxAt5zep55SUdhzKhEdLC.RdVC0saL"&_
'	"&VERSION=2.3&PAYMENTACTION=Authorization&AMT=29.95"&_
'	"&RETURNURL=http://www.blackholenet.com/test/test_paypal_return.asp"&_
'	"&CANCELURL=http://www.blackholenet.com/test/test_paypal_cancel_return.asp"&_
'	"&METHOD=SetExpressCheckout"&_
'	"&RN=2"
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'objHttp.open "POST", "https://api-3t.sandbox.paypal.com/nvp", false
'objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
'objHttp.Send(str)
'Dim objUtil
'Set objUtil = new UtilClass
'response.Write(objUtil.URLDecode(objHttp.responseText)&"<br/><br/>")
'Set objUtil = nothing
'set objHttp = nothing
'************************    FINISH TEST CHIAMATA A PAYPAL SetExpressCheckout E VISUALIZZAZIONE RISPOSTA   ********************




if(Err.number <> 0) then
	response.write("error= ")
	response.write(Err.description)
'else
 'response.write("esecuzione OK")
end if
%>
</body>
</html>