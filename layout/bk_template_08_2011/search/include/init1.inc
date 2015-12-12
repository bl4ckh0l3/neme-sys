<%
Dim Search, objListaAll
Dim strGerarchia, strCSS
Dim totPages, numPage

numPage = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

if(isNull(strGerarchia) OR strGerarchia = "") then strGerarchia = "01"

Set Search = New SearchClass

Dim search_txt
search_txt = request("search_full_txt")
%>