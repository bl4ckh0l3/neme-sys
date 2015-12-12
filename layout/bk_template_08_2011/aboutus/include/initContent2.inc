		<div align="left">		
		<br/>
		<%
		
		if bolHasObj then%>
			<div id="torna"><a href="<%=Application("baseroot") & "/common/include/feedRSS.asp?gerarchia="&strGerarchia%>" target="_blank"><img src="<%=Application("baseroot")&"/common/img/rss_image_medium.gif"%>" vspace="3" hspace="3" border="0" align="right" alt="RSS"></a></div>
			<div>
			<p><strong><%=objCurrentNews.getTitolo()%></strong></p>
			<%if (Len(objCurrentNews.getAbstract1()) > 0) then response.Write(objCurrentNews.getAbstract1()) end if
			if (Len(objCurrentNews.getAbstract2()) > 0) then response.Write(objCurrentNews.getAbstract2()) end if
			if (Len(objCurrentNews.getAbstract3()) > 0) then response.Write(objCurrentNews.getAbstract3()) end if
			response.Write(objCurrentNews.getTesto())%>
			</div>
		<%else
			response.Write("<br/><br/><div align=""center""><strong>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</strong></div>")
		end if%>
		</div>