<%
if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
Set objUserLoggedTmp = nothing
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing

'/**
'* recupero i valori della news selezionata se id_currency <> -1
'*/
Dim id_currency, strDescrizione, iValore, iActive, iDefault, dtInsert, dtRefer
id_currency = request("id_currency")
strDescrizione = ""
iValore = ""
iActive = 0
iDefault = 0
dtInsert = ""
dtRefer = ""

if (Cint(id_currency) <> -1) then
	Dim objCurrency, objSelCurrency
	Set objCurrency = New CurrencyClass
	Set objSelCurrency = objCurrency.findCurrencyByID(id_currency)
	Set objCurrency = nothing
	
	id_currency = objSelCurrency.getID()
	strDescrizione = objSelCurrency.getCurrency()		
	iValore = objSelCurrency.getRate()	
	iActive = objSelCurrency.getActive()	
	iDefault = objSelCurrency.getDefault()	
	dtInsert = objSelCurrency.getDtaInsert()	
	dtRefer = objSelCurrency.getDtaRefer()	
	Set objSelCurrency = Nothing
end if
%>