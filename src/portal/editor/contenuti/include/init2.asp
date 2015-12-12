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
'* recupero i valori della news selezionata se id_news <> -1
'*/
Dim id_news, strTitolo, strAbs, strAbs2, strAbs3, strText, strKeyword, dtData_ins, dtData_pub, dtData_del, stato_news, objTarget, objFiles
Dim page_title, meta_description, meta_keyword, strGeolocal
id_news = request("id_news")
strTitolo = ""
strAbs = ""
strAbs2 = ""
strAbs3 = ""
strText = ""
strKeyword = ""
dtData_ins = ""
dtData_pub = ""
dtData_del = ""
stato_news = ""
page_title = ""
meta_description = ""
meta_keyword = ""
strGeolocal = ""
objTarget = null
objFiles = null

if not (isNull(id_news)) then
	Dim objNews, objSelNews
	Set objNews = New NewsClass
	Set objSelNews = objNews.findNewsByID(id_news)
	Set objNews = nothing

	if not(Instr(1, typename(objSelNews), "NewsClass", 1) > 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")
	end if
	
	id_news = objSelNews.getNewsID()
	strTitolo = objSelNews.getTitolo()
	strAbs = objSelNews.getAbstract1()
	strAbs2 = objSelNews.getAbstract2()
	strAbs3 = objSelNews.getAbstract3()
	strText = objSelNews.getTesto()
	strKeyword = objSelNews.getKeyword()
	dtData_ins = objSelNews.getDataInsNews()
	dtData_pub = objSelNews.getDataPubNews()
	dtData_del = objSelNews.getDataDelNews()
	stato_news = objSelNews.getStato()
	page_title = objSelNews.getPageTitle()
	meta_description = objSelNews.getMetaDescription()
	meta_keyword = objSelNews.getMetaKeyword()
	Set objTarget = objSelNews.getListaTarget()
	
	if not(isNull(objSelNews.getFilePerNews())) then
		Set objFiles = objSelNews.getFilePerNews()	
	end if	

	'********** RECUPERO LA LISTA DI FIELD PRODOTTI ASSOCIATI AL CONTENUTO
	Dim objContentField, objListContentField, hasContentFields
	hasContentFields=false
	On Error Resume Next
	Set objContentField = new ContentFieldClass
	Set objListContentField = objContentField.getListContentField4ContentActive(id_news)
	if(objListContentField.count > 0)then
		hasContentFields=true
	end if
	if(Err.number <> 0) then
		hasContentFields=false
	end if

	'********** RECUPERO LA LISTA DI POINTS GOOGLEMAP ASSOCIATI AL CONTENUTO
	Dim objLocal
	Set objLocal = new LocalizationClass
	On error resume next
	Set points = objLocal.findPointByElement(id_news, 1)
	if not(isNull(points)) then
		for each xLocal in points.Items
			strGeolocal = strGeolocal & langEditor.getTranslated("backend.commons.detail.table.label.latitude") & ", "&langEditor.getTranslated("backend.commons.detail.table.label.longitude")&": "&xLocal.getLatitude()&", "&xLocal.getLongitude()&"<br/>"
		next
	end if
	Set points = nothing	
	if(Err.number <> 0) then	
	end if	
	Set objLocal = nothing
else
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")			
end if
%>