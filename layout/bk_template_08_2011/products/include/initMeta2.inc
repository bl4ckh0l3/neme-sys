<%
on error Resume Next
if(id_prodotto <> "")then
			Set objCurrentProdotto = Prodotto.findProdottoByID(id_prodotto,1)  
end if    

  
if (Instr(1, typename(objCurrentProdotto), "ProductClass", 1) > 0) then
  if not(lang.getTranslated(objCurrentProdotto.getMetaDescription()) = "") then
    metaDescription = metaDescription & " " & lang.getTranslated(objCurrentProdotto.getMetaDescription())
  else
    metaDescription = metaDescription & " " & objCurrentProdotto.getMetaDescription()
  end if
  
  if not(lang.getTranslated(objCurrentProdotto.getMetaKeyword()) = "") then
    metaKeyword = metaKeyword & " " & lang.getTranslated(objCurrentProdotto.getMetaKeyword())
  else
    metaKeyword = metaKeyword & " " & objCurrentProdotto.getMetaKeyword()
  end if
  
  if not(lang.getTranslated(objCurrentProdotto.getPageTitle()) = "") then
    pageTemplateTitle = lang.getTranslated(objCurrentProdotto.getPageTitle())
  else
    if not(objCurrentProdotto.getPageTitle() = "")then
      pageTemplateTitle = objCurrentProdotto.getPageTitle()
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
