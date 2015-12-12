<%
Dim News, objListaNews, objListaTargetCat, objListaTargetLang
Dim strGerarchia, totPages, numPage

Set News = New NewsClass
strGerarchia = request("gerarchia")
if(isNull(strGerarchia) OR strGerarchia = "") then strGerarchia = "01"
numPage = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

Dim stato
stato = 1
if(request("is_preview_content")) then
	stato = null
end if
%>