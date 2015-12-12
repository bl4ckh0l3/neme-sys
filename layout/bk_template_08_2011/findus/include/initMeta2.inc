<%
Dim bolHasObj
bolHasObj = false

on error Resume Next
if(id_news <> "")then
			Set objCurrentNews = News.findNewsByID(id_news)
      bolHasObj = true  
elseif(id_news = "" AND bolCatContainNews AND not(isNull(objListaTargetCat))) then
  Set objListaNews = News.findNews(null, null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_by)	
  
  if(objListaNews.Count > 0) then
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

    arrTmpListNews = objListaNews.Items
    Set objCurrentNews = arrTmpListNews(FromNews)    
  
    bolHasObj = true
  end if
end if    

  
if (Instr(1, typename(objCurrentNews), "NewsClass", 1) > 0) then
  if not(lang.getTranslated(objCurrentNews.getMetaDescription()) = "") then
    metaDescription = metaDescription & " " & lang.getTranslated(objCurrentNews.getMetaDescription())
  else
    metaDescription = metaDescription & " " & objCurrentNews.getMetaDescription()
  end if
  
  if not(lang.getTranslated(objCurrentNews.getMetaKeyword()) = "") then
    metaKeyword = metaKeyword & " " & lang.getTranslated(objCurrentNews.getMetaKeyword())
  else
    metaKeyword = metaKeyword & " " & objCurrentNews.getMetaKeyword()
  end if
  
  if not(lang.getTranslated(objCurrentNews.getPageTitle()) = "") then
    pageTemplateTitle = lang.getTranslated(objCurrentNews.getPageTitle())
  else
    if not(objCurrentNews.getPageTitle() = "")then
      pageTemplateTitle = objCurrentNews.getPageTitle()
    end if
  end if       
end if

metaDescription = Trim(metaDescription)
metaKeyword = Trim(metaKeyword)
pageTemplateTitle = Trim(pageTemplateTitle)

if (pageTemplateTitle = "") then
  pageTemplateTitle = lang.getTranslated("frontend.page.title")
end if      
    
if Err.number <> 0 then
  'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
  bolHasObj = false
end if	%>

<title><%=pageTemplateTitle%></title>
<META name="description" CONTENT="<%=metaDescription%>">
<META name="keywords" CONTENT="<%=metaKeyword%>">

<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
