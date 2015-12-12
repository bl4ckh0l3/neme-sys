<%
Dim News, objListaNews, order_news_by, objListaTargetCat, objListaTargetLang
Dim strGerarchia, totPages, newsXpage, numPage, modelPageNum

Set News = New NewsClass
strGerarchia = request("gerarchia")
order_news_by = 1
newsXpage = 5
numPage = 1
modelPageNum = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

Dim stato
stato = 1
if(request("is_preview_content")) then
	stato = null
end if
%>