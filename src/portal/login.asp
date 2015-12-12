<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001

Dim errorList, strErrorParam, strErrorMessage
Set errorList = Server.CreateObject("Scripting.Dictionary")
errorList.add "001", lang.getTranslated("portal.commons.errors.label.user_in_use")
errorList.add "002", lang.getTranslated("portal.commons.errors.label.no_user_found")
errorList.add "003", lang.getTranslated("portal.commons.errors.label.user_pwd_incorrect")
strErrorMessage = ""
strErrorParam = request("error")
if(errorList.Exists(Trim(LCase(strErrorParam)))) then
	strErrorMessage = "<p><span class=errorText>"&errorList.Item(strErrorParam)&"</span></p>"
end if
Set errorList = Nothing

Dim messageList, strMessageParam
Set messageList = Server.CreateObject("Scripting.Dictionary")
messageList.add "001", lang.getTranslated("portal.commons.messages.label.send_new_password")

strMessageParam = request("message")
if(messageList.Exists(Trim(LCase(strMessageParam)))) then
	strErrorMessage = "<p><span class=errorText>"&messageList.Item(strMessageParam)&"</span></p>"
end if
Set messageList = Nothing
%>

<%
Dim isHTTPS,strLoginAction
isHTTPS = Request.ServerVariables("HTTPS")
If isHTTPS = "off" AND Application("use_https") = 1 Then
	strLoginAction = "https://"&Request.ServerVariables("SERVER_NAME")&Application("baseroot")&"/common/include/VerificaUtente.asp"
Else
	strLoginAction = Application("baseroot")&"/common/include/VerificaUtente.asp"
End If
%>
<!-- #include virtual="/public/layout/include/login.asp" -->