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
	end if%>

  <script>
  function sendPaginationForm(pageNum, actionUrl){
      document.form_pagination.page.value=pageNum;
      document.form_pagination.action=actionUrl;
      document.form_pagination.submit();
  }
  </script>
  <form name="form_pagination" method="post" action="">
  <input type="hidden" name="gerarchia" value="<%=strGerarchia%>">
  <input type="hidden" name="page" value="">
  <%
  listParams = Split(strParameters, "&", -1, 1)
  if(isArray(listParams)) then
    For y=LBound(listParams) to UBound(listParams)
      arrPar = Split(listParams(y), "=", -1, 1)%>
      <input type="hidden" name="<%=arrPar(0)%>" value="<%=arrPar(1)%>">   
    <%next
  end if
  %> 
  </form>  
  
	<%if (page > 1) then%>
	  <a title="<%=lang.getTranslated("portal.commons.pagination.label.prec_page")%>" class="linkPaginazione" href="javascript:sendPaginationForm(<%=(page-1)%>, '<%=sendTopage%>');"><span class="linkPaginazioneLabel"><%=lang.getTranslated("portal.commons.pagination.label.prec_page")%></span>&nbsp;&lt;&nbsp;</a>
	
	<%end if
	
	for i = startPage to EndPage
		if(i = page) then
			class_ = "linkPaginazioneActive"%>
			<a title="<%=lang.getTranslated("portal.commons.pagination.label.page") & " " &  i%>" class="<%=class_%>" href="javascript:sendPaginationForm(<%=(i)%>, '<%=sendTopage%>');"><span class="linkPaginazione">[</span><%=i%><span class="linkPaginazione">]</span></a>
		<%else
			class_ = "linkPaginazione"%>
			<a title="<%=lang.getTranslated("portal.commons.pagination.label.page") & " " &  i%>" class="<%=class_%>" href="javascript:sendPaginationForm(<%=(i)%>, '<%=sendTopage%>');"><%=i%></a>
		<%end if%>	  
	<%next
			
	if (page < totalPages) then%>
	  <a title="<%=lang.getTranslated("portal.commons.pagination.label.next_page")%>" class="linkPaginazione" href="javascript:sendPaginationForm(<%=(page+1)%>, '<%=sendTopage%>');">&nbsp;&gt;&nbsp;<span class="linkPaginazioneLabel"><%=lang.getTranslated("portal.commons.pagination.label.next_page")%></span></a>
	
	<%end if
End Sub
%>