<%
Dim bolHasObj, objCat, objPageTempl
bolHasObj = false

Set objCat = new CategoryClass
Set objPageTempl = new Page4TemplateClass

on error Resume Next
if not(isNull(objListaTargetCat)) then
  Set objListaAll = Search.searchAll(search_txt, objListaTargetLang)
  
  if(objListaAll.Count > 0) then		
    bolHasObj = true
  
    for each t in objListaAll
      if not(lang.getTranslated(objListaAll(t).getMetaDescription()) = "") then
        metaDescription = metaDescription & " " & lang.getTranslated(objListaAll(t).getMetaDescription())
      else
        if not(objListaAll(t).getMetaDescription()="")then
          metaDescription = metaDescription & " " & objListaAll(t).getMetaDescription()
        end if
      end if
      
      if not(lang.getTranslated(objListaAll(t).getMetaKeyword()) = "") then
        metaKeyword = metaKeyword & " " & lang.getTranslated(objListaAll(t).getMetaKeyword())
      else
        if not(objListaAll(t).getMetaKeyword()="")then
          metaKeyword = metaKeyword & " " & objListaAll(t).getMetaKeyword()
        end if
      end if
      
      'if not(lang.getTranslated(objListaAll(t).getPageTitle()) = "") then
      ' pageTemplateTitle = pageTemplateTitle & " " & lang.getTranslated(objListaAll(t).getPageTitle())
      'else
      '  if not(objListaAll(t).getPageTitle()="")then
      '    pageTemplateTitle = pageTemplateTitle & " " & objListaAll(t).getPageTitle()
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
