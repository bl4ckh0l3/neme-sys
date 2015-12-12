<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->

<%
Response.Buffer = TRUE 
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "filename=excel_prod.xls"


Dim objProdotti, objListaProdotti
Set objProdotti = New ProductsClass

Dim TargetTmp, objListaTargetProdTmp, objListaTargetLangTmp, strTargetProd
Set TargetTmp = new TargetClass
objListaTargetProdTmp = null
objListaTargetLangTmp = null


if (not(isNull(request("target_cat"))) AND not(request("target_cat") = "")) then
	Dim objProdTargetTmp, targetX
	
	Set objListaTargetProdTmp = Server.CreateObject("Scripting.Dictionary")
	
	on error resume next
	Set objProdTargetTmp = TargetTmp.findTargetsByCategoria(request("target_cat"))
	for each xProdTargetTmp in objProdTargetTmp.Items
		Set targetX = xProdTargetTmp
		objListaTargetProdTmp.add targetX.getTargetID(), targetX.getTargetDescrizione()
	next
	Set targetX = nothing
	Set objProdTargetTmp = nothing
	
	if Err.number <> 0 then
		objListaTargetProdTmp.add 0, ""
	end if
		
	target_prod_param = strTargetProd
	
	'imposto tutti i target delle lingue per cercare i prodotti
	Dim objTType, idType, objLangTargetTmp
	Set objTType = TargetTmp.getListaTargetType()
	targetLangPrefix = Application("strLangPrefix")
	targetLangPrefix = Replace(targetLangPrefix, "_", "", 1, -1, 1) 
	
	for each x in objTType
		typeDesc = objTType(x)
		if not(InStr(1,typeDesc,targetLangPrefix,0) = 0) then
			idType = x
			Exit For
		end if
	next
	Set objTType = Nothing
	
	Set objLangTargetTmp = TargetTmp.findTargetsByType(idType)
	if not(isNull(objLangTargetTmp)) then
		Set objListaTargetLangTmp = Server.CreateObject("Scripting.Dictionary")
		for each z in objLangTargetTmp
			objListaTargetLangTmp.add objLangTargetTmp(z).getTargetID(), objLangTargetTmp(z).getTargetDescrizione()
		next	
	end if
	Set objLangTargetTmp = Nothing
    
end if
Set TargetTmp = nothing


Dim hasProd
hasProd = false

on error Resume Next
	Set objListaProdotti = objProdotti.findProdotti(null, null, null, null, null, null, null, null, null, order_prod_by, objListaTargetProdTmp, objListaTargetLangTmp, 0, 0)
	
	if(objListaProdotti.Count > 0) then
		hasProd = true
	end if
	
if Err.number <> 0 then
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
	font-weight: bold;
	text-transform:uppercase;
}
</style>
</head>
<body>
<TABLE BORDER=1>
	<tr class="tdHeaderExcel">
		<th><%=langEditor.getTranslated("backend.prodotti.view.table.label.id_prodotto")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.view.table.label.cod_prod")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.view.table.label.nome_prod")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.view.table.label.prod_attivo")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.view.table.label.prezzo_prod")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.view.table.label.tassa_applicata")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.view.table.label.sconto_prod")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.view.table.label.qta_prod")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.lista.table.header.category")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.lista.table.header.lang")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.detail.table.label.prod_download")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.detail.table.label.max_download")%></th>
		<th><%=langEditor.getTranslated("backend.prodotti.detail.table.label.max_download_time")%></th>
	</tr>
<%
if(hasProd) then
	
	Dim prodCounter, iIndex, objTmpProd, objTmpProdKey, objTarget
	iIndex = objListaProdotti.Count	
			
	objTmpProd = objListaProdotti.Items
	objTmpProdKey=objListaProdotti.Keys
	objTarget = null					
			
	for prodCounter = 0 to iIndex-1
		Set objFilteredProd = objTmpProd(prodCounter)
		Set objTarget = objFilteredProd.getListaTarget()
		%>
		<tr>
		<td nowrap align="center"><%=objFilteredProd.getIDProdotto()%></td>
		<td nowrap><%=objFilteredProd.getCodiceProd()%></td>
		<td nowrap><%=objFilteredProd.findFieldTranslation(1 ,langEditor.getLangCode(),1)%></td>
		<td align="center">
		  <%if (objFilteredProd.getAttivo() = 0) then response.Write(langEditor.getTranslated("backend.commons.no"))%>
		  <%if (objFilteredProd.getAttivo() = 1) then response.Write(langEditor.getTranslated("backend.commons.yes"))%>
		</td>
		<td align="right">&euro;&nbsp;<%=FormatNumber(objFilteredProd.getPrezzo(),2,-1)%></td>						
		<td align="center">
		<%Dim objTasse, objTassa
		Set objTasse = new TaxsClass
		Set objTassa = objTasse.findTassaByID(objFilteredProd.getIDTassaApplicata())
		response.write(objTassa.getDescrizioneTassa())
		Set objTassa = nothing	
		Set objTasse = nothing%></td>	
		<td><%=objFilteredProd.getSconto()%>%</td>
		<td><%if(objFilteredProd.getQtaDisp() = Application("unlimited_key"))then%><%=langEditor.getTranslated("backend.prodotti.detail.table.label.qta_unlimited")%><%else%><%=objFilteredProd.getQtaDisp()%><%end if%></td>	
		<td>
		<%	
		Dim CategoriatmpClass, objCategorieXProd	
		if (Instr(1, typename(objTarget), "dictionary", 1) > 0) then
			Set CategoriatmpClass = new CategoryClass
			for each y in objTarget.Keys
				if (objTarget(y).getTargetType() = 2) then
					Set objCategorieXProd = CategoriatmpClass.findCategorieByTargetID(y)
					if not (isNull(objCategorieXProd)) then
						for each z in objCategorieXProd.Keys
							response.write (objCategorieXProd(z).getCatDescrizione() & "<br>")
						next
					end if
					Set objCategorieXProd = nothing
				end if									
			next	
			Set CategoriatmpClass = Nothing
		end if%>						
		</td>
		<td nowrap align="center">
		<%		
		if (Instr(1, typename(objTarget), "dictionary", 1) > 0) then
			for each y in objTarget.Keys
				if (objTarget(y).getTargetType() = 3) then									
					response.write (Replace(objTarget(y).getTargetDescrizione(), "lang_", "", 1, -1, 1) & "<br>")
				end if		
			next		
			Set objTarget = nothing
		end if%>
		</td>						
		<td align="center">
		<%if (objFilteredProd.getProdType() = 0) then response.Write(langEditor.getTranslated("backend.commons.no")) end if%>
		<%if (objFilteredProd.getProdType() = 1) then response.Write(langEditor.getTranslated("backend.commons.yes")) end if%>
		</td>
		<td><%
			Select Case objFilteredProd.getMaxDownload()
			Case -1
				response.Write(langEditor.getTranslated("backend.prodotti.detail.table.label.unlimited"))		
			Case Else
				response.Write(objFilteredProd.getMaxDownload())
			End Select%></td>				  
		<td><%
			Select Case objFilteredProd.getMaxDownloadTime()
			Case -1
				response.Write(langEditor.getTranslated("backend.prodotti.detail.table.label.unlimited"))	
			Case 1
				response.Write("1 "&langEditor.getTranslated("backend.prodotti.detail.table.label.minute"))	
			Case 720
				response.Write("12 "&langEditor.getTranslated("backend.prodotti.detail.table.label.hours"))	
			Case 1440
				response.Write("24 "&langEditor.getTranslated("backend.prodotti.detail.table.label.hours"))		
			Case Else
				response.Write(objFilteredProd.getMaxDownloadTime()&" "&langEditor.getTranslated("backend.prodotti.detail.table.label.minutes"))		
			End Select%></td>
		</tr>				
		<%
		Set objFilteredProd = nothing
	next
	Set objListaProdotti = nothing
				
end if%>

</TABLE>
</body>
</html>
<%
Set objProdotti = Nothing
%>