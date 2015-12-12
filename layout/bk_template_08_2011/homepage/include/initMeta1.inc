<%

Dim bolHasObj
bolHasObj = false

on error Resume Next
if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
  Set objListaNews = News.findNews(null, null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_by)	
  
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
<META http-equiv="content-language" content="<%=LCase(lang.getLangCode())%>">
<META name="owner" CONTENT="Neme-sys - www.neme-sys.org">
<META name="robots" CONTENT="Index, Follow">
<meta name="google-site-verification" content="3rq_o-c-R2qfJRa6Z9tXN5YZak-vi1VljMmLENc60YI" />
