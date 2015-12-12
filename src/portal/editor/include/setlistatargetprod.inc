<%
Dim TargetTmp, objListaTargetProdTmp, objListaTargetLangTmp, strTargetProd
Set TargetTmp = new TargetClass
objListaTargetProdTmp = null
objListaTargetLangTmp = null
strTargetProd = null

if(isNull(session("strTargetProdTmp")) OR session("strTargetProdTmp") = "") OR (not(isNull(request("resetMenu"))) AND request("resetMenu") = "1") then session("strTargetProdTmp") = null end if

if(isNull(request("target_cat")) OR request("target_cat") = "") then 
	strTargetProd = session("strTargetProdTmp") 
else 
	strTargetProd = request("target_cat")
	session("strTargetProdTmp") = strTargetProd
end if

if (not(isNull(strTargetProd)) AND not(strTargetProd = "")) then
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
%>