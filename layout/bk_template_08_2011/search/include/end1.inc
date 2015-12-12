<%
Set objListaAll = nothing
Set Search = nothing

if(Err.number <> 0) then
	'response.Write("test")
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>