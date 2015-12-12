<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->

<%
if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if
	
	Dim id_currency, strDescrizione, iActive, iValore, iDefault, bolDelCurrency, dtaReferer, dtaInsert
	id_currency = request("id_currency")
	strDescrizione = request("descrizione")
	iValore = request("valore")
	iActive = request("attivo")
	iDefault = request("default")
	dtaReferer = request("dta_referer")
	dtaInsert = now()
	bolDelCurrency = request("delete_currency")
	
	Dim objCurrency
	Set objCurrency = New CurrencyClass

	Dim DD, MM, YY, HH, MIN, SS
	
	DD = DatePart("d", dtaReferer)
	MM = DatePart("m", dtaReferer)
	YY = DatePart("yyyy", dtaReferer)	
	dtaReferer = YY&"-"&MM&"-"&DD
	
	DD = DatePart("d", dtaInsert)
	MM = DatePart("m", dtaInsert)
	YY = DatePart("yyyy", dtaInsert)
	HH = DatePart("h", dtaInsert)
	MIN = DatePart("n", dtaInsert)
	SS = DatePart("s", dtaInsert)	
	dtaInsert = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS

	if(iDefault = 1) then
		call objCurrency.resetDefaultCurrency()
	end if

	if (Cint(id_currency) <> -1) then
		if(strComp(bolDelCurrency, "del", 1) = 0) then
			call objCurrency.deleteCurrency(id_currency)
			response.Redirect(Application("baseroot")&"/editor/currency/ListaCurrency.asp")	
		end if
		
		call objCurrency.modifyCurrency(id_currency, strDescrizione, iValore, dtaReferer, dtaInsert, iActive, iDefault)
		Set objCurrency = nothing
		response.Redirect(Application("baseroot")&"/editor/currency/ListaCurrency.asp")		
	else
		call objCurrency.insertCurrency(strDescrizione, iValore, dtaReferer, dtaInsert, iActive, iDefault)
		Set objCurrency = nothing
		response.Redirect(Application("baseroot")&"/editor/currency/ListaCurrency.asp")				
	end if

	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>