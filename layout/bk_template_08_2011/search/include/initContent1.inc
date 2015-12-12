		<%
		'************** codice per la lista e paginazione
		if bolHasObj then%>
			<div id="search">
				<%Dim objSelAll, newsCounter, iIndex, objTmpAll, objTmpKeyAll, FromAll, ToAll, Diff
				Dim splittedGerarchia
				iIndex = objListaAll.Count
				
				response.Write("<p>"&lang.getTranslated("frontend.search.table.label.key_find") & " <strong>" & iIndex &"</strong> " &lang.getTranslated("frontend.search.table.label.key_result_for") & " ""<strong>" & search_txt & "</strong>""</p><br>")
				
				FromAll = ((numPage * elem_x_page) - elem_x_page)
				Diff = (iIndex - ((numPage * elem_x_page)-1))
				if(Diff < 1) then
					Diff = 1
				end if
				
				ToAll = iIndex - Diff
				
				totPages = iIndex\elem_x_page
				if(totPages < 1) then
					totPages = 1
				elseif((iIndex MOD elem_x_page <> 0) AND not ((totPages * elem_x_page) >= iIndex)) then

					totPages = totPages +1	
				end if
						
				objTmpAll = objListaAll.Items
				objTmpKeyAll = objListaAll.Keys
				for newsCounter = FromAll to ToAll
					'per recuperare la gerarchia corretta devo eliminare la parte iniziale con l'id dell news
					'la gerarchia arriva nella forma idnews|gerarchia
					'basta fare una substring che elimini la prima parte 
					splittedGerarchia = objTmpKeyAll(newsCounter)
					splittedGerarchia = Right(splittedGerarchia,(Len(splittedGerarchia)-inStr(splittedGerarchia,"|")))
        
          'recupero la pagina più alta in base al template
          numPageTempl = objPageTempl.getMaxNumPageByIDTemplate(objCat.findExsitingCategoriaByGerarchia(splittedGerarchia).getIDTemplate())
					
					Set objSelAll = objTmpAll(newsCounter)
					if(strComp(typename(objSelAll), "NewsClass", 1) = 0) then%>
						<div>	
						<p>
						<strong><a class="title-ricerca" href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&splittedGerarchia&"&id_news="&objSelAll.getNewsID()&"&page="&numPage&"&modelPageNum="&numPageTempl%>"><%=objSelAll.getTitolo()%></a></strong><br>
						<%=objSelAll.getAbstract1()%>
						</p><p class="line"></p>
						</div>
					<%elseif(strComp(typename(objSelAll), "ProductsClass", 1) = 0) then%>
						<div>	
						<p>
						<strong><a class="title-ricerca" href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&splittedGerarchia&"&id_prodotto="&objSelAll.getIDProdotto()&"&page="&numPage&"&modelPageNum="&numPageTempl%>"><%=objSelAll.getNomeProdotto()%></a></strong><br>
						<%=objSelAll.getSommarioProdotto()%>
						</p><p class="line"></p>
						</div>
					<%end if%>
				<%next
				Set objSelAll = nothing%>			
			    <div>
					<%if(totPages > 1) then%>
						<%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("dir_upload_templ")&"search/Result_search.asp", "&search_full_txt="&search_txt)
					end if%>
			    </div>
			</div>
		<%else
			response.write("<div align=""center""><br/>"&lang.getTranslated("frontend.search.table.label.no_result_found")&"</div>")
		end if%>