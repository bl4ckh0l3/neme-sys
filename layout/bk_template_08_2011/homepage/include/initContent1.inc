		<%
		'************** codice per la lista news e paginazione
		if(bolHasObj) then%>
			<%
			Dim news_counter
			news_counter = 0
			for each x in objListaNews
				Set objSelNews = objListaNews(x)%>
					<h1><%=objSelNews.getTitolo()%></h1>
					<%if (Len(objSelNews.getAbstract1()) > 0) then response.Write("<p>"&objSelNews.getAbstract1()&"</p>") end if%>								
				<%Set objSelNews = nothing
				news_counter = news_counter + 1
			next%>		
		<%end if%>