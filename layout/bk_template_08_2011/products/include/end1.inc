<%
Set objProdField = nothing
Set objProdFieldGroup = nothing
Set currClass = nothing
Set objListaTargetCat = nothing
Set objListaTargetLang = nothing
Set objListaProdotto = nothing
Set Prodotto = Nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>