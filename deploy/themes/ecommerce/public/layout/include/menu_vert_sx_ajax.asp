<!-- #include virtual="/common/include/Objects/DBManagerClass.asp" -->
<!-- #include virtual="/common/include/Objects/CategoryClass.asp" -->
<!-- #include virtual="/common/include/Objects/MenuClass.asp" -->
<!-- #include virtual="/common/include/Objects/LanguageClass.asp" -->
<!-- #include virtual="/common/include/Objects/TemplateClass.asp" -->
<!-- #include virtual="/common/include/Objects/Page4TemplateClass.asp" -->
<!-- #include virtual="/common/include/InitData.inc" -->
<%
Dim menuFruizioneSx, menuCompleteSx, categoriaClassTmpSx
Dim level, iWidth, strSubTmpGer, strSubTmpGerFiltered
Dim iLenGer, iGerLevel, iGerDiff, hrefGer, menuCompleteSxCatLabelTrans

strGerarchia = request("gerarchia")
bolMenuFound = false

Dim x, objCategoriaCheck, strHref
Set menuFruizioneSx = new MenuClass
Set categoriaClassTmpSx = new CategoryClass
Set objTemplateSx = new TemplateClass	
Set objPage4TemplateMenuSx = new Page4TemplateClass   

iGerLevel = menuFruizioneSx.getLivello(strGerarchia)

On Error Resume Next			
Set menuCompleteSx = menuFruizioneSx.getCompleteMenuByMenu("1")
bolMenuFound = true
if(Err.number <>0) then
  bolMenuFound = false
end if
if(bolMenuFound)then
  menuSxCounter = 0
  for each x in menuCompleteSx
    level = menuFruizioneSx.getLivello(x)
    iGerDiff = level - iGerLevel
    menuCompleteSxCatLabelTrans = "frontend.menu.label."&menuCompleteSx(x).getCatDescrizione()
    menuCompleteSxCatDescTrans = "frontend.menu.desc."&menuCompleteSx(x).getCatDescrizione()
  
    if(level > 1) then
      iWidth = ((level-1) * 10)+5 
      strSubTmpGer=x
      if(level>iGerLevel)then
        numDeltaTmpGer = 0
        if(InStrRev(x,".",-1,1)>0)then
          numDeltaTmpGer = Len(x)-(InStrRev(x,".",-1,1)-1)
        end if
        strSubTmpGer = Left(x, Len(x)-numDeltaTmpGer)
      end if
        
      numDeltaSubTmpGer = 0
      if(InStrRev(strSubTmpGer,".",-1,1)>0)then
        numDeltaSubTmpGer = Len(strSubTmpGer)-(InStrRev(strSubTmpGer,".",-1,1)-1)
      end if
      strSubTmpGerFiltered = Left(strSubTmpGer, Len(strSubTmpGer)-numDeltaSubTmpGer)
      
      if(iGerDiff <= 1) then
        if(iGerDiff<=0)then
          strSubTmpGer = strSubTmpGerFiltered
        end if
        if (InStr(1, strGerarchia, strSubTmpGer, 1) > 0) then                
          '*** Controllo se la categoria contiene news, altrimenti cerco la prima sottocategoria che contenga news
          '*** e imposto la nuova gerarchia come parametro nel link
          On Error Resume Next
          Set objCategoriaCheck = categoriaClassTmpSx.checkEmptyCategory(menuCompleteSx(x), true)
          if not(isNull(objCategoriaCheck)) then
            hrefGer = objCategoriaCheck.getCatGerarchia()
            Set objTemplateSelectedSx = objTemplateSx.findTemplateByID(objCategoriaCheck.findLangTemplateXCategoria(lang.getLangCode(),true))
            strHref = menuFruizioneSx.resolveHrefUrl(base_url, 1, lang, objCategoriaCheck, objTemplateSelectedSx, objPage4TemplateMenuSx)
            Set objTemplateSelectedSx = nothing
          else
            strHref = "#"                  
          end if
          Set objCategoriaCheck = nothing
          if(Err.number <>0) then
            strHref = "#"
          end if

          '*** checkSelectedCategory
          bolSelectedCat = false
          strSubSelCat = strGerarchia
          for a=1 to Abs(iGerDiff)
              strSubSelCat = Left(strSubSelCat,InStrRev(strSubSelCat,".",-1,1)-1)
          next
          
          if(strComp(x, strSubSelCat, 1) = 0) then
              bolSelectedCat = true
          end if              
          %>
          <li><a href="javascript:openLinkMenuSX('<%=hrefGer%>','<%=strHref%>');" style="padding-left:<%=iWidth%>px;" <%if(bolSelectedCat) then response.Write("class=""link-attivo-menu-sub""") else response.Write("class=""link-menu-sub""") end if%>><%if not(isNull(lang.getTranslated(menuCompleteSxCatLabelTrans))) AND not(lang.getTranslated(menuCompleteSxCatLabelTrans) = "") then response.write(lang.getTranslated(menuCompleteSxCatLabelTrans)) else response.Write(menuCompleteSx(x).getCatDescrizione()) end if%></a>
          <%if not(isNull(lang.getTranslated(menuCompleteSxCatDescTrans))) AND not(lang.getTranslated(menuCompleteSxCatDescTrans) = "") then%><p style="padding-left:<%=iWidth+5%>px;"><%=lang.getTranslated(menuCompleteSxCatDescTrans)%></p><%end if%>
          </li>		
        <%end if
      end if
    else
      iWidth = 0

      strSubTmpGer = strGerarchia
      numDeltaTmpGer = 0
      if(InStr(1, strGerarchia, ".", 1) > 0)then
        numDeltaTmpGer = Len(strGerarchia)-(InStr(1, strGerarchia, ".", 1)-1)
      end if
      strSubTmpGer = Left(strGerarchia, Len(strGerarchia)-numDeltaTmpGer)
    
      '*** Controllo se la categoria contiene news, altrimenti cerco la prima sottocategoria che contenga news
      '*** e imposto la nuova gerarchia come parametro nel link
      On Error Resume Next
      Set objCategoriaCheck = categoriaClassTmpSx.checkEmptyCategory(menuCompleteSx(x), true)
      if not(isNull(objCategoriaCheck)) then
        hrefGer = objCategoriaCheck.getCatGerarchia()
        Set objTemplateSelectedSx = objTemplateSx.findTemplateByID(objCategoriaCheck.findLangTemplateXCategoria(lang.getLangCode(),true))
        strHref = menuFruizioneSx.resolveHrefUrl(base_url, 1, lang, objCategoriaCheck, objTemplateSelectedSx, objPage4TemplateMenuSx)
        Set objTemplateSelectedSx = nothing
      else
        strHref = "#"                  
      end if
      Set objCategoriaCheck = nothing
      if(Err.number <>0) then
        strHref = "#"
      end if%>
      <li><a href="javascript:openLinkMenuSX('<%=hrefGer%>','<%=strHref%>');" <%if(strComp(x, strSubTmpGer, 1) = 0) then response.Write("class=""link-attivo""")%>><%if not(isNull(lang.getTranslated(menuCompleteSxCatLabelTrans))) AND not(lang.getTranslated(menuCompleteSxCatLabelTrans) = "") then response.write(lang.getTranslated(menuCompleteSxCatLabelTrans)) else response.Write(menuCompleteSx(x).getCatDescrizione()) end if%></a>
      <%if not(isNull(lang.getTranslated(menuCompleteSxCatDescTrans))) AND not(lang.getTranslated(menuCompleteSxCatDescTrans) = "") then%><p style="padding-left:<%=iWidth+5%>px;"><%=lang.getTranslated(menuCompleteSxCatDescTrans)%></p><%end if%></li>
    <%end if
    menuSxCounter = menuSxCounter +1
  next
end if
Set objPage4TemplateMenuSx = nothing
Set objTemplateSx = nothing
Set categoriaClassTmpSx = nothing
Set menuCompleteSx = nothing
Set menuFruizioneSx = nothing%>