<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->

<%
Response.Buffer = TRUE 
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "filename=excel_order.xls"

Dim objOrdini, objListaOrdini, objListaStatiOrdine
Dim objPayment, objTmpPayment, objListaPayment, objUtente
Set objOrdini = New OrderClass
Set objPayment = New PaymentClass
Set objUtente = New UserClass

Dim hasOrder, id_ordine_search
hasOrder = false

if(request("search_ordini") = "") then
	on error Resume Next
	Set objListaOrdini = objOrdini.getListaOrdini(order_ordine_by, 0)
	Set objListaStatiOrdine = objOrdini.getListaStatiOrder()		

	if(objListaOrdini.Count > 0) then
		hasOrder = true
	end if

	if Err.number <> 0 then
	end if		
else		
	id_ordine_search = null
	
	id_user_search = request("id_utente_search")
	if(id_user_search = "") then id_user_search = null end if
	dta_ins_search_from = request("dta_ins_search_from")
	if(dta_ins_search_from = "") then dta_ins_search_from = null end if
	dta_ins_search_to = request("dta_ins_search_to")
	if(dta_ins_search_to = "") then dta_ins_search_to = null end if
	tipo_pagam_search = request("tipo_pagam_search")
	if(tipo_pagam_search = "") then tipo_pagam_search = null end if
	pagam_done_search = request("pagam_done_search")
	if(pagam_done_search = "") then pagam_done_search = null end if
	stato_ord_search = request("stato_ord_search")
	if(stato_ord_search = "") then stato_ord_search = null end if
	ord_by_search = request("ord_by_search")
	if(ord_by_search = "") then ord_by_search = null end if
	ord_guid_search = request("ord_guid_search")
	if(ord_guid_search = "") then ord_guid_search = null end if		

	on error Resume Next
	Set objListaOrdini = objOrdini.findOrdini(id_ordine_search, id_user_search, dta_ins_search_from, dta_ins_search_to, stato_ord_search, tipo_pagam_search, pagam_done_search, ord_by_search, 0, ord_guid_search)
	Set objListaStatiOrdine = objOrdini.getListaStatiOrder()		

	if(objListaOrdini.Count > 0) then
		hasOrder = true
	end if
		
	if Err.number <> 0 then
	end if	
end if
%>

<html>
<head>
<title></title>
<style type="text/css"> 
body {
	background: #FFFFFF;
}
.tdHeaderExcel {
	background-color: #432D30;
	text-align: left;
	color: #FFFFFF;
	font-weight:bold;
}
</style>
</head>
<body>
<TABLE BORDER=1>
	<tr class="tdHeaderExcel">
		<td><%=langEditor.getTranslated("backend.ordini.lista.table.header.id_ordine")%></td>
		<td><%=langEditor.getTranslated("backend.ordini.lista.table.header.cliente")%></td>
		<td><%=langEditor.getTranslated("backend.ordini.lista.table.search.header.data_insert")%></td>
		<td><%=langEditor.getTranslated("backend.ordini.lista.table.search.header.type_pagam")%></td>
		<td><%=langEditor.getTranslated("backend.ordini.lista.table.search.header.pagam_done")%></td>
		<td><%=langEditor.getTranslated("backend.ordini.lista.table.header.totale_order")%></td>
		<td><%=langEditor.getTranslated("backend.ordini.lista.table.search.header.stato_ord")%></td>
	</tr>
<%
if(hasOrder) then
	
	intCount = 0					
	iIndex = objListaOrdini.Count
	iIndexStatiOrder = objListaStatiOrdine.Count		
			
	objTmpOrder = objListaOrdini.Items					
	Dim objTmpUser
								
	for counter = 0 to iIndex-1
		Set objFilteredOrder = objTmpOrder(counter)%>
		<tr>			
		<td><%=objFilteredOrder.getIDOrdine()%></td>
		<td>
		<%
		Set objTmpUser = objUtente.findUserByID(objFilteredOrder.getIDUtente())
		response.Write(objTmpUser.getUserName())
		%>
		</td>
		<td><%=objFilteredOrder.getDtaInserimento()%></td>
		<td>
		<%
		Set objTmpPayment = objPayment.findPaymentByID(objFilteredOrder.getTipoPagam())
		payUrlTmp = objTmpPayment.getURL()
		response.write(langEditor.getTranslated(objTmpPayment.getKeywordMultilingua()))
		Set objTmpPayment = Nothing
		%>
		</td>
		<td>
		<%
		Select Case objFilteredOrder.getPagamEffettuato()
		Case 0%>
			&nbsp;&nbsp;<img src="<%=Application("baseroot")&"/editor/img/sema_al.png"%>" alt="<%=langEditor.getTranslated("backend.ordini.lista.table.alt.order_to_pay")%>" hspace="2" vspace="0" border="0" align="absmiddle">
		<%	response.write(langEditor.getTranslated("backend.commons.no"))						
		Case 1						
			Dim paymentNotified
			paymentNotified = false
			Set objPaymentTrans = new PaymentTransactionClass
			if(objPaymentTrans.hasPaymentTransactionNotified(objFilteredOrder.getIDOrdine())) then
				paymentNotified = true
			end if
			Set objPaymentTrans = nothing
			
			if(paymentNotified OR payUrlTmp = 0) then%>
				&nbsp;&nbsp;<img src="<%=Application("baseroot")&"/editor/img/sema_no.png"%>" alt="<%=langEditor.getTranslated("backend.ordini.lista.table.alt.order_paied_notified")%>" hspace="2" vspace="0" border="0" align="absmiddle">
			<%else%>
				&nbsp;&nbsp;<img src="<%=Application("baseroot")&"/editor/img/sema_adup.png"%>" alt="<%=langEditor.getTranslated("backend.ordini.lista.table.alt.order_paied_no_notified")%>" hspace="2" vspace="0" border="0" align="absmiddle">
			<%end if							
			response.write(langEditor.getTranslated("backend.commons.yes"))
		Case Else
		End Select
		%>
		</td>
		<td>&euro;&nbsp;<%=FormatNumber(objFilteredOrder.getTotale(),2,-1)%></td>
		<td>
			<%=langEditor.getTranslated(objListaStatiOrdine(objFilteredOrder.getStatoOrdine()))%>
		</td>
		</tr>				
		<%intCount = intCount +1
		Set objFilteredOrder = nothing
	next
	Set objTmpUser = nothing
	Set objTmpOrder = nothing
	Set objTmpStatiOrder = nothing
	Set objTmpStatiOrderKey = nothing
	Set objListaOrdini = nothing
				
end if%>

</TABLE>
</body>
</html>
<%
Set objPayment = Nothing
Set objOrdini = Nothing	
Set objUtente = nothing	
%>