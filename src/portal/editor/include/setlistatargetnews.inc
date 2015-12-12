<%
Dim TargetTmp, objListaTargetCatTmp, objListaTargetLangTmp, strTargetCat
Set TargetTmp = new TargetClass		
objListaTargetCatTmp = null
objListaTargetLangTmp = null
strTargetCat = null

if(isNull(session("strTargetCatTmp")) OR session("strTargetCatTmp") = "") OR (not(isNull(request("resetMenu"))) AND request("resetMenu") = "1") then session("strTargetCatTmp") = null end if

if(isNull(request("target_cat")) OR request("target_cat") = "") then 
	strTargetCat = session("strTargetCatTmp") 
else 
	strTargetCat = request("target_cat")
	session("strTargetCatTmp") = strTargetCat
end if

if (not(isNull(strTargetCat)) AND not(strTargetCat = "")) then
	Dim objCatTargetTmp, targetX
	
	Set objListaTargetCatTmp = Server.CreateObject("Scripting.Dictionary")
	
	on error resume next
	Set objCatTargetTmp = TargetTmp.findTargetsByCategoria(strTargetCat)
	for each xCatTargetTmp in objCatTargetTmp.Items
		Set targetX = xCatTargetTmp
		objListaTargetCatTmp.add targetX.getTargetID(), targetX.getTargetDescrizione()
	next
	Set targetX = nothing
	Set objCatTargetTmp = nothing

	if Err.number <> 0 then
		objListaTargetCatTmp.add 0, ""
	end if
	
	target_cat_param = strTargetCat
	
	'imposto tutti i target delle lingue per cercare le news
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