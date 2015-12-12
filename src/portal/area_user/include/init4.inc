<%
if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
Set objUserLoggedTmp = nothing

'/**
'* recupero i valori della news selezionata se id_prod <> -1
'*/
Dim id_order, id_utente, dta_ins, totale_ord, tipo_pagam, pagam_done, stato_order, objSelProdPerOrder
id_order = request("id_ordine")
id_utente = ""
dta_ins = ""
totale_ord = 0
tipo_pagam = ""
payment_commission = 0
pagam_done = 0
stato_order = 0
objSelProdPerOrder = null
sconto_cliente = objUserLogged.getSconto()

Dim objUtente, objTmpUser
					
if not (isNull(id_order)) then
	Dim objOrdini, objSelOrdine, objProdPerOrder
	Set objOrdini = New OrderClass
	Set objSelOrdine = objOrdini.findOrdineByID(id_order, 1)
	Set objProdPerOrder = New Products4OrderClass
	Set objOrdini = nothing
	
	id_order = objSelOrdine.getIDOrdine()
	id_utente = objSelOrdine.getIDUtente()
	dta_ins = objSelOrdine.getDtaInserimento()
	totale_ord = objSelOrdine.getTotale()
	tipo_pagam = objSelOrdine.getTipoPagam()
	payment_commission = objSelOrdine.getPaymentCommission()
	pagam_done = objSelOrdine.getPagamEffettuato()
	stato_order = objSelOrdine.getStatoOrdine()
	
	if (isObject(objProdPerOrder.getListaProdottiXOrdine(id_order)) AND not(isNull(objProdPerOrder.getListaProdottiXOrdine(id_order))) AND not(isEmpty(objProdPerOrder.getListaProdottiXOrdine(id_order)))) then
		Set objSelProdPerOrder = objProdPerOrder.getListaProdottiXOrdine(id_order)	
	end if
		
	Set objUtente = New UserClass
else
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")			
end if

Set objProdField = new ProductFieldClass

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>