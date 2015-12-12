		<div align="left" id="contenuti">
		<%
		'************** codice per la lista news e paginazione		
		
		if(bolHasObj) then%>
				<br/>	
        <div id="content-center-prodotto">
          <%Dim objSelNews, newsCounter, iIndex, objTmpNews, FromNews, ToNews, Diff
          iIndex = objListaNews.Count
          FromNews = ((numPage * elem_x_page) - elem_x_page)
          Diff = (iIndex - ((numPage * elem_x_page)-1))
          if(Diff < 1) then
            Diff = 1
          end if
          
          ToNews = iIndex - Diff
          
          totPages = iIndex\elem_x_page
          if(totPages < 1) then
            totPages = 1
          elseif((iIndex MOD elem_x_page <> 0) AND not ((totPages * elem_x_page) >= iIndex)) then
            totPages = totPages +1	
          end if		
              
          objTmpNews = objListaNews.Items		
          for newsCounter = FromNews to ToNews
          'for newsCounter = 0 to objListaNews.Count -1
            Set objSelNews = objTmpNews(newsCounter)%>
            
            <div id="prodotto-immagine">
              <%if not(isNull(objSelNews.getFilePerNews())) AND not(isEmpty(objSelNews.getFilePerNews())) then%>
                <%Dim hasNotSmallImg
                hasNotSmallImg = true
                Set objListaFilePerNews = objSelNews.getFilePerNews()			
                for each xObjFile in objListaFilePerNews
                  Set objFileXNews = objListaFilePerNews(xObjFile)
                  iTypeFile = objFileXNews.getFileTypeLabel()
                  if(Cint(iTypeFile) = 1) then%>	
                    <img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" alt="<%=objSelNews.getTitolo()%>" width="140" height="130" />
                    <%hasNotSmallImg = false
                    Exit for
                  end if
                  Set objFileXNews = nothing	
                next		
                if(hasNotSmallImg) then%>
                <img width="140" height="130" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0">
                <%end if
                Set objListaFilePerNews = nothing
                else%>
                <img width="140" height="130" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0">
                <%end if%>          
            </div>
                
            <div id="prodotto-testo"><p class="title_contenuti"><a href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&id_news="&objSelNews.getNewsID()&"&page="&numPage&"&modelPageNum="&(modelPageNum+1)%>"><%=objSelNews.getTitolo()%></a></p>
            <%if (Len(objSelNews.getAbstract1()) > 0) then response.Write(objSelNews.getAbstract1()) end if%>
            </div>
            <div id="clear"></div>
            <div id="prodotto-footer"></div>
            <%Set objSelNews = nothing
          next%>
            <div>
          <%if(totPages > 1) then%>
            <%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("controller_page"), "")
          end if%>
          </div>
				<div id="torna"><a href="<%=Application("baseroot") & "/common/include/feedRSS.asp?gerarchia="&strGerarchia%>" target="_blank"><img src="<%=Application("baseroot")&"/common/img/rss_image.gif"%>" vspace="3" hspace="3" border="0" align="right" alt="RSS"></a></div>
			  </div>       
		<%else
			response.Write("<br/><br/><div align=""center""><strong>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</strong></div>")
		end if%>
		</div>