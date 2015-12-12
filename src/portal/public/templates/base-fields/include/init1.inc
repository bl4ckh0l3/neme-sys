<%
Dim News, objListaNews, objListaTargetCat, objListaTargetLang
Dim strGerarchia, totPages, numPage, stato

Set News = New NewsClass
strGerarchia = request("gerarchia")
numPage = 1
stato = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

if(request("is_preview_content")) then
	stato = null
end if
%>
<!-- #include virtual="/common/include/setTemplateTargetList.inc" -->
<%
Dim bolHasObj
bolHasObj = false

Set objContentField = new ContentFieldClass

if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
  on error Resume Next
  Set objListaNews = News.findNewsSlimCached(null, null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_by,false,false)  
  if(objListaNews.Count > 0) then		
    bolHasObj = true
  end if
  
  if Err.number <> 0 then
    bolHasObj = false
  end if
  
  if(bolHasObj) then
    Set objListPairKeyValue = Server.CreateObject("Scripting.Dictionary")	
    bolHasFieldsFilter = false
    strParamPagFilter=""
    strValParFilter =""

    On Error Resume Next
  
    '****************** verifico se è stata fatta una ricerca dei field per contenuto e in tal caso recupero solo le news filtrate
    if(request("fields_filter")="1")then 
      if (Request.Form.Count>0)then      
        For Each y In Request.Form
          if (Instr(1, y, "field_", 1) > 0) then
            if(Trim(Request.Form(y))<>"") then
              tmpValue = Split(Request.Form(y), "#", -1, 1)
              objListPairKeyValue.add tmpValue(0), tmpValue(1)
              strParamPagFilter=strParamPagFilter&y&"="&Request.Form(y)&"&"
              strValParFilter =strValParFilter & Request.Form(y)&"|"
            end if
          end if
        Next 

        if(objListPairKeyValue.Count > 0)then
          bolHasFieldsFilter = true
          strParamPagFilter=strParamPagFilter&"fields_filter=1"
          Session("fields_filter") = Left(strValParFilter, Len(strValParFilter)-1)
          Session("fields_filter_param") = strParamPagFilter
        end if
      end if 
    elseif(Len(Session("fields_filter")) > 0 AND request("reset_fields_filter")<>"1")then
      tmpFVal=Split(Session("fields_filter"), "|", -1, 1)
      for counter=0 to Ubound(tmpFVal)
          tmpSubValue = Split(tmpFVal(counter), "#", -1, 1)
          objListPairKeyValue.add tmpSubValue(0), tmpSubValue(1)
      next
      strParamPagFilter = Session("fields_filter_param")
      bolHasFieldsFilter = true
    end if
    if(request("reset_fields_filter")="1")then
      Session.Contents.Remove("fields_filter")
      Session.Contents.Remove("fields_filter_param")
      Set objListPairKeyValue = Server.CreateObject("Scripting.Dictionary")	
    end if

    if(Err.number <> 0) then
      'response.write(Err.description)
      bolHasFieldsFilter = false
    end if    
  
    for each t in objListaNews
      On Error Resume Next
      keepContent=true      
      Set objListContentFieldTmp = objContentField.getListContentField4ContentActive(t)      
      if(objListContentFieldTmp.Count > 0)then
        objListaNews(t).setListaFields(objListContentFieldTmp)
      
        if(bolHasFieldsFilter)then 
          for each j in objListPairKeyValue
            if(objListContentFieldTmp.Exists(Clng(j))=false OR strComp(objListContentFieldTmp(Clng(j)).getSelValue(), objListPairKeyValue(j), 1) <> 0)then
              objListaNews.remove(t)
              keepContent=false
              exit for
            end if
          next  
        end if
      end if
      
      if(Err.number <> 0) then
        keepContent = true
      end if
      Set objListContentFieldTmp = nothing
    
      if(keepContent)then
        if not(lang.getTranslated(objListaNews(t).getMetaDescription()) = "") then
          metaDescription = metaDescription & " " & lang.getTranslated(objListaNews(t).getMetaDescription())
        else
          if not(objListaNews(t).getMetaDescription()="")then
            metaDescription = metaDescription & " " & objListaNews(t).getMetaDescription()
          end if
        end if
        
        if not(lang.getTranslated(objListaNews(t).getMetaKeyword()) = "") then
          metaKeyword = metaKeyword & " " & lang.getTranslated(objListaNews(t).getMetaKeyword())
        else
          if not(objListaNews(t).getMetaKeyword()="")then
            metaKeyword = metaKeyword & " " & objListaNews(t).getMetaKeyword()
          end if
        end if 
      end if
    next

    metaDescription = Trim(metaDescription)
    metaKeyword = Trim(metaKeyword)
      
    Dim objSelNews, newsCounter, iIndex, objTmpNews, FromNews, ToNews, Diff
    iIndex = objListaNews.Count
    FromNews = ((numPage * elem_x_page) - elem_x_page)
    Diff = (iIndex - ((numPage * elem_x_page)-1))
    if(Diff < 1) then
      Diff = 1
    end if

    ToNews = iIndex - Diff

    totPages = iIndex\elem_x_page
    if(totPages < 1) then
      totPages = 1
    elseif((iIndex MOD elem_x_page <> 0) AND not ((totPages * elem_x_page) >= iIndex)) then
      totPages = totPages +1	
    end if		
        
    objTmpNews = objListaNews.Items     
  end if
end if

if (pageTemplateTitle = "") then
  pageTemplateTitle = lang.getTranslated("frontend.page.title")
end if  

Response.Charset="UTF-8"
Session.CodePage  = 65001

Set objCat = new CategoryClass
Set objPageTempl = new Page4TemplateClass
Set objTemplate = new TemplateClass
Set objMenuFruizione = new MenuClass
bolHasDetailLink = false
On Error Resume Next							         
Set objCategoriaTmp = objCat.findExsitingCategoriaByGerarchia(strGerarchia)
if not(isNull(objCategoriaTmp)) then
  Set objTemplateSelected = objTemplate.findTemplateByID(objCategoriaTmp.findLangTemplateXCategoria(lang.getLangCode(),true))
  bolHasDetailLink = true
else
  bolHasDetailLink = false                  
end if
if(Err.number <>0) then
  bolHasDetailLink = false
end if 
%>