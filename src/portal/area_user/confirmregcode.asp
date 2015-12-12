<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001

Dim id_user, registration_code, confirmed
id_user = request("id_user")
registration_code = request("confirm_registration_code")

Dim objUtente
Set objUtente = New UserClass
confirmed = objUtente.findConfirmationCodeUserByID(id_user,registration_code)%>
<!-- #include virtual="/public/layout/area_user/confirmregcode.asp" -->
<%Set objUtente = nothing%>
