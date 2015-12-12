<%
if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp?from="&request.ServerVariables("URL"))
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
Set objUserLoggedTmp = nothing

Dim objOrdini, objListaOrdini, objListaStatiOrdine
Dim objPayment, objTmpPayment, objListaPayment
Set objOrdini = New OrderClass
Set objPayment = New PaymentClass

Dim order_ordine_by, reqOrderBy
order_ordine_by = 6
reqOrderBy = request("order_by")

if (not(isNull(reqOrderBy)) AND not(reqOrderBy = "")) then
	order_ordine_by = reqOrderBy	
end if

Dim totPages, orderXpage, numPage

orderXpage = 10
numPage = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if


Dim statiOrderCount, iIndexStatiOrder, objTmpStatiOrder, objTmpStatiOrderKey
Dim orderCounter, iIndex, objTmpOrder, objTmpOrderKey, FromOrder, ToOrder, Diff
Dim objUtente, objListaUtenti, userCounter, objTmpUtenti, objTmpUtentiKey, tmpObjUsr
Dim intCount, styleRow, styleRow2

Response.Charset="UTF-8"
Session.CodePage  = 65001
		
Dim id_ordine_search, id_utente_search, dta_ins_search_from, dta_ins_search_to, tipo_pagam_search, pagam_done_search, stato_ord_search
id_ordine_search = null
id_utente_search = objUserLogged.getUserID()
dta_ins_search_from = null
dta_ins_search_to = null
tipo_pagam_search = null
pagam_done_search = null
stato_ord_search = null

if(id_utente_search = "") then id_utente_search = null end if
if(dta_ins_search_from = "") then dta_ins_search = null end if
if(dta_ins_search_to = "") then dta_ins_search = null end if
if(tipo_pagam_search = "") then tipo_pagam_search = null end if
if(pagam_done_search = "") then pagam_done_search = null end if
if(stato_ord_search = "") then stato_ord_search = null end if
if(ord_by_search = "") then ord_by_search = null end if

Dim hasOrder
hasOrder = false
on error Resume Next
	Set objListaOrdini = objOrdini.findOrdini(id_ordine_search, id_utente_search, dta_ins_search_from, dta_ins_search_to, stato_ord_search, tipo_pagam_search, pagam_done_search, order_ordine_by, 0, null)
	Set objListaStatiOrdine = objOrdini.getListaStatiOrder()		

	if(objListaOrdini.Count > 0) then
		hasOrder = true
	end if
	
if Err.number <> 0 then
end if	
%>