<%
Dim Categoria, objCategoria, Target, objLangTarget, bolCatContainNews, bolCatContainProd, pageURL, modelPageNum, templateDir
Dim metaDescription, metaKeyword, pageTemplateTitle
Dim order_by, elem_x_page
bolCatContainNews = false
bolCatContainProd = false

Set Categoria = new CategoryClass
Set Target = new TargetClass
Set objTemplate = new TemplateClass

'************** VERIFICO SE ESISTE LA GERARCHIA, IN CASO CONTRARIO LA ESTRAGGO DAL TEMPLATE
'************** MECCANISMO DIFENSIVO SU CHIAMATA A PAGINA DIRETTA SENZA PARAMETRI (GOOGLE)
if(strGerarchia="")then
  ' recupero url corrente per definire path
  tmpurl = request.ServerVariables("URL")
  tmpurl = Mid(tmpurl,1,InStrRev(tmpurl,"/",-1,1)) 
  tmpurl = Mid(tmpurl,1,InStrRev(tmpurl,"/"&lang.getLangCode()&"/",-1,1)-1) 
  tmpurl = right(tmpurl,Len(tmpurl)-InStrRev(tmpurl,"/",-1,1))
  strGerarchia = Categoria.findGerarchiaByTemplateDirectory(tmpurl)
end if

Set objCategoria = Categoria.findCategoriaByGerarchia(strGerarchia)
catDescription = objCategoria.getCatDescrizione()
bolCatContainNews = objCategoria.contieneNews()
bolCatContainProd = objCategoria.contieneProd()

metaDescription = ""
metaKeyword = ""
pageTemplateTitle = ""

if not(lang.getTranslated(objCategoria.getMetaDescription()) = "") then
  metaDescription = lang.getTranslated(objCategoria.getMetaDescription())
else
  metaDescription = objCategoria.getMetaDescription()
end if

if not(lang.getTranslated(objCategoria.getMetaKeyword()) = "") then
  metaKeyword = lang.getTranslated(objCategoria.getMetaKeyword())
else
  metaKeyword = objCategoria.getMetaKeyword()
end if

if not(lang.getTranslated(objCategoria.getpageTitle()) = "") then
  pageTemplateTitle = lang.getTranslated(objCategoria.getpageTitle())
else
  pageTemplateTitle = objCategoria.getpageTitle()
end if


objListaTargetCat = null
if not(isNull(objCategoria.getTargetPerCategoria(objCategoria.getCatID()))) then
	Set objListaTargetCat = objCategoria.getTargetPerCategoria(objCategoria.getCatID())
end if

Set objListaTargetLang = Server.CreateObject("Scripting.Dictionary")
Set objLangTarget = Target.findTargetByDescNoTransaction(Application("strLangPrefix") & lang.getLangCode())
' mi salvo l'id del target lingua da usare per i filtri dei contenuti
langIdTarget = objLangTarget.getTargetID()
objListaTargetLang.add objLangTarget.getTargetID(), objLangTarget.getTargetDescrizione()
Set objLangTarget = nothing
Set Target = nothing
'imposto il css specifico di questa categoria
Set objTmpTemplate = objTemplate.findTemplateByID(objCategoria.getIDTemplate())

if not(objTmpTemplate.getTemplateCssWithPath()="") then
  strCSS = Application("dir_upload_templ")&objTmpTemplate.getTemplateCssWithPath()
end if

'imposto in numero di pagine per template, l'ordinamento e il numero di pagina del template selezionato
order_by = objTmpTemplate.getOrderBy()

if not(isNull(request("order_by"))) AND not(request("order_by") = "") then
	order_by = request("order_by")
end if

elem_x_page = objTmpTemplate.getElemXPage()

Set objP4T = new Page4TemplateClass
pageURL = request.ServerVariables("URL")
pageURL = Mid(pageURL,InStrRev(pageURL,"/",-1,1)+1)
modelPageNum = objP4T.findPagePerTemplateByFileName(objCategoria.getIDTemplate(),pageURL).getPageNum()
templateDir = objTmpTemplate.getDirTemplate()

Set objP4T = nothing
Set objTmpTemplate = nothing
Set objTemplate = nothing
Set objCategoria = nothing
Set Categoria = Nothing
%>