<%
Sub PaginazioneFrontend(totPages, currPage, strGerarchia, sendTopage, strParameters)
	page = Cint(currPage)
	totalPages = Cint(totPages)
	Max = 10
	
	EndPage = Cint(page + max)
	startPage = 1
	
	' qualche controllo
	if (EndPage> totalPages) then
		EndPage = totalPages
	end if

	' qualche controllo
	if (startPage < page) then
		startPage = page
	end if
	if(Cbool((EndPage - startPage) < Max)) then
		startPage = (EndPage - Max)
	end if
	if (startPage < 1) then
		startPage = 1
	end if

	if (page > 1) then%>
	  <a title="<%=langEditor.getTranslated("portal.commons.pagination.label.prec_page")%>" class="link-paginazione" href="<%=Application("baseroot") & sendTopage &"?gerarchia="&strGerarchia&"&page="&(page-1)&strParameters%>"><span class="link-paginazione"><%=langEditor.getTranslated("portal.commons.pagination.label.prec_page")%></span></a>
	
	<%end if
	
	for i = startPage to EndPage
		if(i = page) then
			class_ = "link-paginazione-active"
		else
			class_ = "link-paginazione"
		end if%>
	  <a title="<%=langEditor.getTranslated("portal.commons.pagination.label.page") & " " &  i%>" class="<%=class_%>" href="<%=Application("baseroot") & sendTopage &"?gerarchia="&strGerarchia&"&page="&i&strParameters%>"><span class="link-paginazione">[</span><%=i%><span class="link-paginazione">]</span></a>
	<%next
			
	if (page < totalPages) then%>
	  <a title="<%=langEditor.getTranslated("portal.commons.pagination.label.next_page")%>" class="link-paginazione" href="<%=Application("baseroot") & sendTopage &"?gerarchia="&strGerarchia&"&page="&(page+1)&strParameters%>"><span class="link-paginazione"><%=langEditor.getTranslated("portal.commons.pagination.label.next_page")%></span></a>
	
	<%end if
End Sub
%>