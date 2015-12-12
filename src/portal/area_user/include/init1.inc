<%
if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp, objListaRuoli
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objListaRuoli = objUserLoggedTmp.getListaRuoli()
	Set objUserLoggedTmp = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp?from=area_user")
end if

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>