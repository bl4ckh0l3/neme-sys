<%
Dim Search, objListaNews
Dim strGerarchia, strCSS
Dim totPages, numPage

numPage = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

strGerarchia = request("gerarchia")
if(isNull(strGerarchia) OR strGerarchia = "") then strGerarchia = "01"

order_news_by = 1

Set Search = New SearchClass

Dim titolo, abstract, abstract2, abstract3, text, keyword, search_txt
search_txt = request("search_full_txt")

titolo = search_txt
abstract = search_txt
abstract2 = search_txt
abstract3 = search_txt
text = search_txt
keyword = search_txt
%>