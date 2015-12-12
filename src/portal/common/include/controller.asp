<!-- #include file="Objects/DBManagerClass.asp" -->
<!-- #include file="Objects/CategoryClass.asp" -->
<!-- #include file="Objects/Page4TemplateClass.asp" -->
<!-- #include file="Objects/TemplateClass.asp" -->
<!-- #include file="Objects/LanguageClass.asp" -->
<!-- #include file="Objects/UserClass.asp" -->
<!-- #include file="Objects/MenuClass.asp" -->
<!-- #include file="InitData.inc" -->
<%
Dim Categoria, selCat, reqGerarchia
Dim objTemplate, objTmpTemplate, objPage4Template, pageRedirect, modelPageNum, pageRedirectDefault
Dim isSubdomainActive, langcodeDir
Set objTemplate = new TemplateClass
Set objPage4Template = new Page4TemplateClass
Set objCategoria = new CategoryClass
Set objMenuFruizione = new MenuClass
modelPageNum = 1
reqGerarchia = Trim(request("gerarchia"))
if not(request("modelPageNum") = "") then
	modelPageNum = request("modelPageNum")
end if

pageRedirectDefault = Application("baseroot") & "/default.asp?empty=1"
bolAddGer = false
On Error Resume Next
if(reqGerarchia="" OR reqGerarchia="00") then
	Set selCat = objCategoria.findFirstAvailableCategoria()
	reqGerarchia = selCat.getCatGerarchia()
	bolAddGer = true
else
	Set selCat = objCategoria.findExsitingCategoriaByGerarchia(reqGerarchia)
end if
Set objCategoriaCheck = objCategoria.checkEmptyCategory(selCat, true)
Set selCat = nothing

if not(isNull(objCategoriaCheck)) then
	Set objTemplateSelected = objTemplate.findTemplateByID(objCategoriaCheck.findLangTemplateXCategoria(lang.getLangCode(),true))
	pageRedirect = objMenuFruizione.resolveHrefUrl(base_url, modelPageNum, lang, objCategoriaCheck, objTemplateSelected, objPage4Template)
	if(strComp(Trim(pageRedirect), "#", 1) = 0)then 
		pageRedirect = pageRedirectDefault 
	end if
	Set objTemplateSelected = nothing
else
  pageRedirect = pageRedirectDefault                  
end if
Set objCategoriaCheck = nothing
if(Err.number <>0) then
	pageRedirect = pageRedirectDefault
end if
Set objMenuFruizione = nothing
Set objPage4Template = nothing
Set objTemplate = nothing
Set objCategoria = nothing

'***************************************************************************************************************************************
'*******************************************************************  REDIRECT  ******************************************************
'***************************************************************************************************************************************

'*** metodo con 301 per evitare problemi di indicizzazione con google
'Response.Status="301 Moved Permanently"

'*** metodo indiretto con FORM
Dim ArrQueryStringParams, objListPairKeyValue
Set objListPairKeyValue = Server.CreateObject("Scripting.Dictionary")

if (Request.QueryString.Count>0)then
	Set objListPairKeyValue = Request.QueryString
elseif (Request.Form.Count>0)then
	Set objListPairKeyValue = Request.Form
end if
if(bolAddGer)then
	objListPairKeyValue.item("gerarchia") = reqGerarchia
end if
%>
<HTML>
<BODY onload="document.controller_redirect.submit();">
<form method="post" name="controller_redirect" action="<%=pageRedirect%>">
<%For Each y In objListPairKeyValue%>
<input type="hidden" name="<%=y%>" value="<%=objListPairKeyValue(y)%>">
<%Next
Set objListPairKeyValue = nothing%>
</form>
</BODY>
</HTML>