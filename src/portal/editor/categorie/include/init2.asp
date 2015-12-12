<%
if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp, objListaTargetPerUser
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))

Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if

if not(isNull(objUserLogged.getTargetPerUser(objUserLogged.getUserID()))) then
	Set objListaTargetPerUser = objUserLogged.getTargetPerUser(objUserLogged.getUserID())	
end if
Set objUserLoggedTmp = nothing
Set objUserLogged = nothing

'/**
'* recupero i valori della news selezionata se id_target <> -1
'*/
Dim id_categoria, iNumMenu, iGerarchia, strDescrizione, catType, bolContNews, bolContProd, bolVisible
Dim objTargets, objCatTarget, metaDescription, metaKeyword, pageTitle, sub_domain_url
Dim idTemplate
id_categoria = request("id_categoria")
iNumMenu = ""
iGerarchia = "" 
strDescrizione = ""
catType = ""
bolContNews = false
bolContProd = false
bolVisible = true
idTemplate = -1
metaDescription = ""
metaKeyword = ""
pageTitle = ""
sub_domain_url = ""
objCatTarget = null

Dim objCategoria, objSelCategoria
Set objCategoria = new CategoryClass

if (Cint(id_categoria) <> -1) then
	Set objSelCategoria = objCategoria.findCategoriaByID(id_categoria)
	
	id_categoria = objSelCategoria.getCatID()
	iNumMenu = objSelCategoria.getNumMenu()
	iGerarchia = objSelCategoria.getCatGerarchia()
	strDescrizione = objSelCategoria.getCatDescrizione()
	catType = objSelCategoria.getCatType()
	bolContNews = objSelCategoria.contieneNews()
	bolContProd = objSelCategoria.contieneProd()
	bolVisible = objSelCategoria.isCatVisible()
	idTemplate = objSelCategoria.getIDTemplate()
	metaDescription = objSelCategoria.getMetaDescription()
	metaKeyword = objSelCategoria.getMetaKeyword()
	pageTitle = objSelCategoria.getPageTitle()
	sub_domain_url = objSelCategoria.getSubDomainURL()

	if not(isNull(objSelCategoria.getTargetPerCategoria(id_categoria))) then
	Set objCatTarget = objSelCategoria.getTargetPerCategoria(id_categoria)
	end if			
end if
%>