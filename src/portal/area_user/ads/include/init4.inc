<%
if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp, objListaTargetPerUser,id_utente
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
id_utente = Session("objUtenteLogged")
if not(strComp(Cint(strRuoloLogged), Application("guest_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLoggedTmp = nothing

'/**
'* recupero i valori della news selezionata se id_news <> -1
'*/
Dim id_ads, id_element, ads_type, price, strPhone, dta_ins
id_element = request("id_news")
id_ads = -1
ads_type = 0
price = 0
strPhone = ""
dta_ins = now()
%>