<%
if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
Set objUserLoggedTmp = nothing
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("guest_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing

Dim objNews, objListaNews
Set objNews = New NewsClass

Dim order_news_by

if not(request("order_by") = "") then
	session("order_news_by") = request("order_by")
	order_news_by = session("order_news_by")
else
	if not(session("order_news_by") = "") then
		order_news_by = session("order_news_by")
	else
		session("order_news_by") = 1
		order_news_by = session("order_news_by")
	end if
end if

Dim totPages, itemsXpage, numPage

if not(request("items") = "") then
	session("contenutiItems") = request("items")
	itemsXpage = session("contenutiItems")
	session("contenutiPage") = 1
else
	if not(session("contenutiItems") = "") then
		itemsXpage = session("contenutiItems")
	else
		session("contenutiItems") = 20
		itemsXpage = session("contenutiItems")
	end if
end if

if not(request("page") = "") then
	session("contenutiPage") = request("page")
	numPage = session("contenutiPage")
else
	if not(session("contenutiPage") = "") then
		numPage = session("contenutiPage")
	else
		session("contenutiPage") = 1
		numPage = session("contenutiPage")
	end if
end if

Dim target_cat_param
target_cat_param = ""

Dim CategoriatmpClass, objListCatXNews
Set CategoriatmpClass = new CategoryClass

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