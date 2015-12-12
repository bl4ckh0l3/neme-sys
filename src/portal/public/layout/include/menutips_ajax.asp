<!-- #include virtual="/common/include/Objects/DBManagerClass.asp" -->
<!-- #include virtual="/common/include/Objects/CategoryClass.asp" -->
<!-- #include virtual="/common/include/Objects/MenuClass.asp" -->
<!-- #include virtual="/common/include/Objects/LanguageClass.asp" -->
<!-- #include virtual="/common/include/Objects/TemplateClass.asp" -->
<!-- #include virtual="/common/include/Objects/Page4TemplateClass.asp" -->
<!-- #include virtual="/common/include/InitData.inc" -->
<%
Dim menuFruizioneTips, menuTips, gerTmp, menuTipsCatDescTrans
Set menuFruizioneTips = new MenuClass
Set categoriaClassTips = new CategoryClass
Set objTemplateTips = new TemplateClass	
Set objPage4TemplateMenuTips = new Page4TemplateClass   
if(not(isNull(request("gerarchia"))) AND not(request("gerarchia") = "")) then
	Set menuTips = menuFruizioneTips.getTipsMenu(request("gerarchia"))
	
	Dim q, menuTipsCounter
	menuTipsCounter = 0 
	for each q in menuTips
		gerTmp = menuTips(q).getCatDescrizione()
    menuTipsCatDescTrans = "frontend.menu.label."&gerTmp			
		
		'*** Controllo se la categoria contiene news, altrimenti cerco la prima sottocategoria che contenga news
		'*** e imposto la nuova gerarchia come parametro nel link
    On Error Resume Next
    Set objCategoriaCheckTips = categoriaClassTips.checkEmptyCategory(menuTips(q), true)
    if not(isNull(objCategoriaCheckTips)) then    
      hrefGerTips = objCategoriaCheckTips.getCatGerarchia()
      Set objTemplateSelectedTips = objTemplateTips.findTemplateByID(objCategoriaCheckTips.findLangTemplateXCategoria(lang.getLangCode(),true))
      strHrefTips = menuFruizioneTips.resolveHrefUrl(base_url, 1, lang, objCategoriaCheckTips, objTemplateSelectedTips, objPage4TemplateMenuTips)
      Set objTemplateSelectedTips = nothing
    else
      strHrefTips = "#"                  
    end if    
		Set objCategoriaCheckTips = nothing
    if(Err.number <>0) then
      strHrefTips = "#"
    end if%>
		<a href="javascript:sendMenuTips('<%=hrefGerTips%>','<%=strHrefTips%>');"><%if not(isNull(lang.getTranslated(menuTipsCatDescTrans))) AND not(lang.getTranslated(menuTipsCatDescTrans) = "") then response.write(lang.getTranslated(menuTipsCatDescTrans)) else response.Write(gerTmp) end if%></a> <%if(menuTipsCounter<UBound(menuTips.Keys)) then response.write("-->") end if%>
	<%	menuTipsCounter = menuTipsCounter + 1
	next
	Set menuTips = nothing
end if
Set objPage4TemplateMenuTips = nothing
Set objTemplateTips = nothing
Set categoriaClassTips = nothing
Set menuFruizioneTips = nothing%>