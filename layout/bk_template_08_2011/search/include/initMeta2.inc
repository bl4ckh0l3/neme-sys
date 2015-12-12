<%
Dim bolHasObj, objCat, objPageTempl
bolHasObj = false

Set objCat = new CategoryClass
Set objPageTempl = new Page4TemplateClass

on error Resume Next
if not(isNull(objListaTargetCat)) then
  Set objListaNews = Search.searchNews(null, titolo, abstract, abstract2, abstract3, text, keyword, null, objListaTargetLang, null, null, 1, order_by, false)
  
  if(objListaNews.Count > 0) then		
    bolHasObj = true
  
    for each t in objListaNews
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
      
      'if not(lang.getTranslated(objListaNews(t).getPageTitle()) = "") then
      ' pageTemplateTitle = pageTemplateTitle & " " & lang.getTranslated(objListaNews(t).getPageTitle())
      'else
      '  if not(objListaNews(t).getPageTitle()="")then
      '    pageTemplateTitle = pageTemplateTitle & " " & objListaNews(t).getPageTitle()
      '  end if
      'end if       
    next

    metaDescription = Trim(metaDescription)
    metaKeyword = Trim(metaKeyword)
    'pageTemplateTitle = Trim(pageTemplateTitle)  
  end if
end if
  
if Err.number <> 0 then
  bolHasObj = false
end if

if (pageTemplateTitle = "") then
  pageTemplateTitle = lang.getTranslated("frontend.page.title")
end if  %>

<title><%=pageTemplateTitle%></title>
<META name="description" CONTENT="<%=metaDescription%>">
<META name="keywords" CONTENT="<%=metaKeyword%>">

<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
