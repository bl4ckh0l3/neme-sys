<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
Dim id_news, News, objCurrentNews, strCSS, objFileXNews, objListaFilePerNews

Set News = New NewsClass
id_news = request("id_news")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<link rel="stylesheet" href="<%="http://" & request.ServerVariables("SERVER_NAME") & Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<link rel="stylesheet" href="<%="http://" & request.ServerVariables("SERVER_NAME") & Application("baseroot") &Application("dir_upload_templ")&"newsletter/css/templateNewsletter.css"%>" type="text/css">
<SCRIPT SRC="<%="http://" & request.ServerVariables("SERVER_NAME") & Application("baseroot") & "/common/js/javascript_global.js"%>"></SCRIPT>
</head>
<body>
<table class="tableContainer" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td class="tdContainerTop">&nbsp;</td>
  </tr>
  <tr>
    <td class="tdContainerContent">
	<table class="tableContent" border="0" align="left" cellpadding="0" cellspacing="0">
	  <tr>
		<td class="tdContent" align="center"><img src="<%="http://" & request.ServerVariables("SERVER_NAME") & Application("baseroot")&Application("dir_upload_templ")&"newsletter/img/newsletter_1.jpg"%>" vspace="0" hspace="0" border="0"></td>
	  </tr>
	  <tr>
		<td class="tdContent">
		<%Set objCurrentNews = News.findNewsByID(id_news)
		
		response.Write("<span class=""titoloNews"">"&objCurrentNews.getTitolo() & "</span><br><br>")
		if (Len(objCurrentNews.getAbstract1()) > 0) then response.Write(objCurrentNews.getAbstract1() & "<br>") end if
		if (Len(objCurrentNews.getAbstract2()) > 0) then response.Write(objCurrentNews.getAbstract2() & "<br>") end if
		if (Len(objCurrentNews.getAbstract3()) > 0) then response.Write(objCurrentNews.getAbstract3() & "<br><br>") end if
		response.Write(objCurrentNews.getTesto() & "<br><br>")
		Set objCurrentNews = nothing
		%>
		</td>
	  </tr>
	</table>
	</td>
  </tr>
  <tr>
    <td class="tdContainerBott">
	</td>
  </tr>
</table>
</body>
</html>
<%
Set News = Nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>
