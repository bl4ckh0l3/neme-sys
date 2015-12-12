<?xml version="1.0"?>
<%On Error Resume Next

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
' Impostazione che setta il tipo di file in output su XML
response.ContentType = "text/xml"

Dim News, objListaNews, objListaTargetCat, objListaTargetLang
Dim strGerarchia, newsXpage, numPage, order_news_by

Set News = New NewsClass
strGerarchia = request("gerarchia")
order_news_by = 1
newsXpage = 20
numPage = 1
stato = 1

pageRssURL = ""
Dim isHTTPS
isHTTPS = Request.ServerVariables("HTTPS")
If isHTTPS = "off" AND Application("use_https") = 1 Then
	pageRssURL = "https://"&Request.ServerVariables("SERVER_NAME")& "/common/include/Controller.asp?gerarchia="&strGerarchia
Else
	pageRssURL = "http://"&Request.ServerVariables("SERVER_NAME")& "/common/include/Controller.asp?gerarchia="&strGerarchia
End If
%>
<!-- #include virtual="/common/include/setTemplateTargetList.inc" -->
<rss version="2.0">
  <channel> 
    <title>Feed RSS: <%=catDescription%></title> 
    <link><%=Request.ServerVariables("SERVER_NAME")%></link> 
    <description>Feed RSS: <%=catDescription%></description> 
    <language><%=lang.getLangCode()%></language> 
	<%
	'************** codice per la lista news e paginazione
	Dim bolHasObj
	bolHasObj = false
	
	on error Resume Next
	if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
		Set objListaNews = News.findNewsSlimCached(null, null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_news_by,false,false)	
		
		if(objListaNews.Count > 0) then		
			bolHasObj = true
		end if
	end if
		
	if Err.number <> 0 then
		bolHasObj = false
	end if			
	
	if(bolHasObj) then
		
		for each x in objListaNews
			Set objSelNews = objListaNews(x)%>
			<item>
			<title><![CDATA[<%=objSelNews.getTitolo()%>]]></title>
			<description><![CDATA[
			<%if (Len(objSelNews.getAbstract1()) > 0) then response.Write(objSelNews.getAbstract1() & "<br>") end if
			if (Len(objSelNews.getAbstract2()) > 0) then response.Write(objSelNews.getAbstract2() & "<br>") end if
			if (Len(objSelNews.getAbstract3()) > 0) then response.Write(objSelNews.getAbstract3() & "<br>") end if%>
			]]></description>
			<link><![CDATA[<%=pageRssURL&"&id_news="&objSelNews.getNewsID()&"&page="&numPage&"&modelPageNum="&(modelPageNum+1)%>]]></link>
			</item>
			<%Set objSelNews = nothing
		next		
	end if%>
  </channel> 
</rss>

<%
Set objListaTargetCat = nothing
Set objListaTargetLang = nothing
Set objListaNews = nothing
Set News = Nothing

if(Err.number <> 0) then
	response.write(Err.description)
end if
%> 