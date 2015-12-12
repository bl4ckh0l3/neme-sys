<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<%
Dim url, objHttp

On Error Resume Next

url = "http://"&Application("srt_default_server_name")&Application("baseroot")&"/editor/currency/currencyPoller.asp"
		
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
objHttp.open "POST", url, false
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send()
set objHttp = nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
else
	response.Redirect(Application("baseroot")&"/editor/currency/ListaCurrency.asp")
end if
%>
