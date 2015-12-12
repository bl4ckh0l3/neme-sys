<%
Dim bolHasObj
bolHasObj = false

on error Resume Next
if (bolCatContainProd) AND not(isNull(objListaTargetCat)) then
  Set objListaProdotto = Prodotto.findProdotti(null, null, null, null, null, null, null, 0, order_by, objListaTargetCat, objListaTargetLang, 1)	
  
  if(objListaProdotto.Count > 0) then		
    bolHasObj = true
  
    for each t in objListaProdotto
      if not(lang.getTranslated(objListaProdotto(t).getMetaDescription()) = "") then
        metaDescription = metaDescription & " " & lang.getTranslated(objListaProdotto(t).getMetaDescription())
      else
        if not(objListaProdotto(t).getMetaDescription()="")then
          metaDescription = metaDescription & " " & objListaProdotto(t).getMetaDescription()
        end if
      end if
      
      if not(lang.getTranslated(objListaProdotto(t).getMetaKeyword()) = "") then
        metaKeyword = metaKeyword & " " & lang.getTranslated(objListaProdotto(t).getMetaKeyword())
      else
        if not(objListaProdotto(t).getMetaKeyword()="")then
          metaKeyword = metaKeyword & " " & objListaProdotto(t).getMetaKeyword()
        end if
      end if
      
      'if not(lang.getTranslated(objListaProdotto(t).getPageTitle()) = "") then
      ' pageTemplateTitle = pageTemplateTitle & " " & lang.getTranslated(objListaProdotto(t).getPageTitle())
      'else
      '  if not(objListaProdotto(t).getPageTitle()="")then
      '    pageTemplateTitle = pageTemplateTitle & " " & objListaProdotto(t).getPageTitle()
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
<meta http-equiv="content-language" content="<%=LCase(lang.getLangCode())%>">
