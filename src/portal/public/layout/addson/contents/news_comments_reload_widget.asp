<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
id_news = ""
if(request("id_news")<>"1") then
	id_news=request("id_news")
end if%>
<!-- #include file="news_comments_widget.asp" -->