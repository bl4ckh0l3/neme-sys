<%
'*** metodo indiretto con FORM (se si usa QUERY_STRING commentare tutta la parte di codice seguente, compreso HTML)
Dim ArrQueryStringParams, objListPairKeyValue
ArrQueryStringParams = Split(Session("payment_opt"),"&",-1,1)
Set Session("payment_opt") = nothing
Set objListPairKeyValue = Server.CreateObject("Scripting.Dictionary")

For Each x In ArrQueryStringParams
	key = Left(x,InStr(1,x,"=",1)-1)
	value =  Right(x,(Len(x)-InStrRev(x,"=",-1,1)))	
	objListPairKeyValue.add key, value
Next

pageRedirect = Application("baseroot") & "/test/test_paypal_return.asp"
%>
<HTML>
<BODY onload="document.controller_redirect.submit();">
<form method="post" name="controller_redirect" action="<%=pageRedirect%>">
<%For Each y In objListPairKeyValue%>
<input type="hidden" name="<%=y%>" value="<%=objListPairKeyValue(y)%>">
<%Next
Set objListPairKeyValue = nothing%>
</form>
</BODY>
</HTML>