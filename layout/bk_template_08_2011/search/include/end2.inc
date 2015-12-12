<%
Set objListaNews = nothing
Set Search = nothing
'Set objListaTarget = nothing
'Set News = Nothing
'Set Categoria = Nothing

if(Err.number <> 0) then
	'response.Write("test")
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>