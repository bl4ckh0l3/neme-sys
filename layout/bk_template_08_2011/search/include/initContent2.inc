		<%
		'************** codice per la lista news e paginazione
		if bolHasObj then%>
			<div id="search">
				<%Dim objSelNews, newsCounter, iIndex, objTmpNews, objTmpKeyNews, FromNews, ToNews, Diff, numPageTempl
				Dim splittedGerarchia
				iIndex = objListaNews.Count%>
				<p>				
				<%=lang.getTranslated("frontend.search.table.label.key_find") & " <strong>" & iIndex &"</strong> " &lang.getTranslated("frontend.search.table.label.key_result_for") & " ""<strong>" & search_txt & "</strong>""<br><br><br>"%>
				</p>
				<%FromNews = ((numPage * elem_x_page) - elem_x_page)
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
				objTmpKeyNews = objListaNews.Keys
				for newsCounter = FromNews to ToNews
					'per recuperare la gerarchia corretta devo eliminare la parte iniziale con l'id dell news
					'la gerarchia arriva nella forma idnews|gerarchia
					'basta fare una substring che elimini la prima parte 
					splittedGerarchia = objTmpKeyNews(newsCounter)
					splittedGerarchia = Right(splittedGerarchia,(Len(splittedGerarchia)-inStr(splittedGerarchia,"|")))
        
          'recupero la pagina più alta in base al template
          numPageTempl = objPageTempl.getMaxNumPageByIDTemplate(objCat.findExsitingCategoriaByGerarchia(splittedGerarchia).getIDTemplate())
					
					Set objSelNews = objTmpNews(newsCounter)%>
					<div>	
						<p><strong><a class="title-ricerca" href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&splittedGerarchia&"&id_news="&objSelNews.getNewsID()&"&page="&numPage&"&modelPageNum="&numPageTempl%>"><%=objSelNews.getTitolo()%></a></strong><br>
						<%if not(objSelNews.getAbstract1() = "") then response.write(objSelNews.getAbstract1()) end if%>
						</p><p class="line"></p>
					</div>
					<%Set objSelNews = nothing
				next%>
			  	<div>
				<%if(totPages > 1) then%>
					<%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("dir_upload_templ")&"search/Result_search_news.asp", "&search_full_txt="&search_txt)
				end if%>
				</div>
			</div>
		<%else
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&objListaNews.getException())
			response.write("<div align=""center""><br/>"&lang.getTranslated("frontend.search.table.label.no_result_found")&"</div>")
		end if
    
    Set objPageTempl = nothing
    Set objCat = nothing
    %>