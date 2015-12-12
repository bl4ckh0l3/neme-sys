
		<%
		'************** codice per la lista Prodotto e paginazione	
    
		if(bolHasObj) then		
			Dim objSelProdotto, objListaFile, ProdottoCounter, iIndex, objTmpProdotto, FromProdotto, ToProdotto, Diff
			iIndex = objListaProdotto.Count
			FromProdotto = ((numPage * elem_x_page) - elem_x_page)
			Diff = (iIndex - ((numPage * elem_x_page)-1))
			if(Diff < 1) then
				Diff = 1
			end if
			
			ToProdotto = iIndex - Diff
			
			totPages = iIndex\elem_x_page
			if(totPages < 1) then
				totPages = 1
			elseif((iIndex MOD elem_x_page <> 0) AND not ((totPages * elem_x_page) >= iIndex)) then
				totPages = totPages +1	
			end if		
					
			objTmpProdotto = objListaProdotto.Items%>

			<div id="content-center-prodotto">
				<div align="left">
        <%
        Dim order_by_price, order_by_name
        if(request("order_by")=105) then 
          order_by_price= 106
          order_by_name = 103
        end if
        if(request("order_by")=103) then 
          order_by_price= 105
          order_by_name = 104
        end if
        
        Select Case request("order_by")
        Case 105
          order_by_price= 106
          order_by_name = 103
        Case 106
          order_by_price= 105
          order_by_name = 103
        Case 103
          order_by_price= 105
          order_by_name = 104
        Case 104
          order_by_price= 105
          order_by_name = 103
        Case Else
          order_by_price= 105
          order_by_name = 103
				End Select
        %>
        <a href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&page="&numPage&"&order_by="&order_by_price&"&modelPageNum="&modelPageNum%>"><%=lang.getTranslated("frontend.template_prodotto.table.label.orderby_price")%></a>
        &nbsp;|&nbsp;<a href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&page="&numPage&"&order_by="&order_by_name&"&modelPageNum="&modelPageNum%>"><%=lang.getTranslated("frontend.template_prodotto.table.label.orderby_name")%></a>
        <%if(Application("disable_ecommerce") = 0) OR (Application("disable_ecommerce") = 1) then%>
          <%On Error Resume Next
          Set currencyList = currClass.getListaCurrency(null, 1, null)'%>
          &nbsp;|&nbsp;<span class="prodotto_select_label"><%=lang.getTranslated("frontend.template.currency.label")%></span>&nbsp;<select class="prodotto_select" name="currency" onchange="javascript:changeCurrency(this);">
          <option value=""></option>
          <%for each x in currencyList%>
            <option value="<%=currencyList(x).getCurrency()%>" <%if(currencyList(x).getCurrency() = Session("currency")) then response.write("selected") end if%>><%=currencyList(x).getCurrency()%> (<%=lang.getTranslated("backend.currency.keyword.label."&currencyList(x).getCurrency())%>)</option>
          <%next%>
          </select>
          <%
          Set currencyList = nothing
          if(Err.number <> 0) then 
            'response.write(Err.description) 
          end if%>
        <%end if%>
				<br/><br/>
				</div>
				<%
				Set objTasse = new TaxsClass
				for ProdottoCounter = FromProdotto to ToProdotto
					Set objSelProdotto = objTmpProdotto(ProdottoCounter)
					Set objListaFile = New File4ProductsClass%>
					<div>
					<form action="<%=Application("baseroot")&Application("dir_upload_templ")&"shopping-card/ManageCarrello.asp"%>" method="post" name="form_add_to_carrello_<%=ProdottoCounter%>" enctype="multipart/form-data"><!--  onSubmit="return addToCarrello(document.form_add_to_carrello_<%=ProdottoCounter%>,<%=ProdottoCounter%>);" -->	
					<input type="hidden" value="<%=objSelProdotto.getIDProdotto()%>" name="id_prodotto">
					<input type="hidden" value="add" name="operation">						
					<input type="hidden" value="<%=ProdottoCounter%>" name="form_counter">		
					<div id="prodotto-immagine">
					<%if not(isNull(objListaFile.getFilePerProdotto(objSelProdotto.getIDProdotto()))) AND not(isEmpty(objListaFile.getFilePerProdotto(objSelProdotto.getIDProdotto()))) then%>
						<%Dim hasNotSmallImg
						hasNotSmallImg = true
						Set objListaFilePerProdotto = objListaFile.getFilePerProdotto(objSelProdotto.getIDProdotto())					
						for each xObjFile in objListaFilePerProdotto
							Set objFileXProdotto = objListaFilePerProdotto(xObjFile)
							iTypeFile = objFileXProdotto.getFileTypeLabel()
							if(Cint(iTypeFile) = 1) then%>	
								<img src="<%=Application("dir_upload_prod")&objFileXProdotto.getFilePath()%>" alt="<%=objSelProdotto.getNomeProdotto()%>" width="140" height="130" />
								<%hasNotSmallImg = false
								Exit for
							end if
							Set objFileXProdotto = nothing	
						next		
						if(hasNotSmallImg) then%>
						<img width="140" height="130" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0">
						<%end if
						Set objListaFilePerProdotto = nothing
						else%>
						<img width="140" height="130" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0">
						<%end if%>
					</div>				
					<div id="prodotto-testo">
						<h2><a href="<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&id_prodotto="&objSelProdotto.getIDProdotto()&"&page="&numPage&"&modelPageNum="&(modelPageNum+1)%>"><%=objSelProdotto.findFieldTranslation(1 ,lang.getLangCode(),1)%></a></h2>
						<p><%=objSelProdotto.findFieldTranslation(2 ,lang.getLangCode(),1)%></p>
						<%if(Application("disable_ecommerce") = 0) OR (Application("disable_ecommerce") = 1) then%>
            <p>
            
						<%           
						Dim numPrezzoReal, descTassa, numPrezzoOld
						numPrezzoReal = objSelProdotto.getPrezzo() 
            numPrezzoOld = numPrezzoReal 
            discountPercent = 0
            
            if(hasGroup) then
              On Error Resume Next
							discountPercent = objSelMargin.getDiscountPercentual(CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objSelProdotto.getsconto()),CDbl(scontoCliente))
              numPrezzoReal = objSelMargin.getAmount(numPrezzoReal,CDbl(objSelMargin.getMargin()),CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objSelProdotto.getsconto()),CDbl(scontoCliente))
              if(Err.number <>0) then
              end if
            else
              if(objSelProdotto.hasSconto() AND (not(hasSconto) OR (hasSconto AND Application("manage_sconti") = 1))) then
                numPrezzoReal = objSelProdotto.getPrezzoScontato()
							  discountPercent = CDbl(objSelProdotto.getsconto())
							  if(hasSconto)then
                  numPrezzoReal = numPrezzoReal - (numPrezzoReal / 100 * scontoCliente)
                  discountPercent = discountPercent+CDbl(scontoCliente)
							  end if
              else
                if(hasSconto)then
                numPrezzoReal = numPrezzoReal - (numPrezzoReal / 100 * scontoCliente)
                discountPercent = CDbl(scontoCliente)
                end if
              end if
            end if%>
            
						<input type="hidden" value="<%=numPrezzoReal%>" name="prezzo">
						<%
            
            '***********************************   INTERNAZIONALIZZAZIONE TASSE   ****************************
						descTassa = ""
            applyOrigTax = true
            if(Application("enable_international_tax_option")=1) AND (international_country_code<>"") then
              if(hasGroup AND (Instr(1, typename(groupClienteTax), "TaxsGroupClass", 1) > 0)) then
                On Error Resume Next
                ' verifico se l'utente ha selezionato il flag tipologia cliente=società e se per il country/region selezionato il falg escludi tassa è attivo
                if(Cint(userIsCompanyClient)=1 AND groupClienteTax.isTaxExclusion(groupClienteTax.getID(), international_country_code,international_state_region_code))then
                  descTassa = lang.getTranslated("frontend.prodotti.label.tax_excluded")							
                  applyOrigTax = false
                else
                  objRelatedTax = groupClienteTax.findRelatedTax(groupClienteTax.getID(), international_country_code,international_state_region_code)
                  if(not(isNull(objRelatedTax))) then
                    Set objTaxG = objTasse.findTassaByID(objRelatedTax)
                    numPrezzoReal = numPrezzoReal+groupClienteTax.getImportoTassa(numPrezzoReal, objTaxG)
                    numPrezzoOld = numPrezzoOld+groupClienteTax.getImportoTassa(numPrezzoOld, objTaxG)
                    descTassa = objTaxG.getDescrizioneTassa()
                    Set objTaxG = nothing
                    applyOrigTax = false
                  else
                    applyOrigTax = true		
                  end if
                end if                
                if(Err.number<>0)then
                  applyOrigTax = true
                end if	
              else
                On Error Resume Next
                Set groupProdTax = objSelProdotto.getTaxGroupObj(objSelProdotto.getTaxGroup()) 
                if(Instr(1, typename(groupProdTax), "TaxsGroupClass", 1) > 0) then
                  ' verifico se l'utente ha selezionato il flag tipologia cliente=società e se per il country/region selezionato il falg escludi tassa è attivo
                  if(Cint(userIsCompanyClient)=1 AND groupProdTax.isTaxExclusion(groupProdTax.getID(), international_country_code,international_state_region_code))then
                    descTassa = lang.getTranslated("frontend.prodotti.label.tax_excluded")							
                    applyOrigTax = false
                  else	
                    objRelatedTax = groupProdTax.findRelatedTax(groupProdTax.getID(), international_country_code,international_state_region_code)			  
                    if(not(isNull(objRelatedTax))) then
                      Set objTaxG = objTasse.findTassaByID(objRelatedTax)
                      numPrezzoReal = numPrezzoReal+groupProdTax.getImportoTassa(numPrezzoReal, objTaxG)
                      numPrezzoOld = numPrezzoOld+groupProdTax.getImportoTassa(numPrezzoOld, objTaxG)
                      descTassa = objTaxG.getDescrizioneTassa()
                      Set objTaxG = nothing
                      applyOrigTax = false		
                    end if
                  end if                  
                else
                  applyOrigTax = true
                end if
                Set groupProdTax = nothing
                if(Err.number<>0)then
                  applyOrigTax = true
                end if
              end if
            end if
            if(applyOrigTax)then
              descTassa = ""
              if not(isNull(objSelProdotto.getIDTassaApplicata())) AND not(objSelProdotto.getIDTassaApplicata() = "") then
                numPrezzoReal = numPrezzoReal+objSelProdotto.getImportoTassa(numPrezzoReal)
                numPrezzoOld = numPrezzoOld+objSelProdotto.getImportoTassa(numPrezzoOld)
                descTassa = objTasse.findTassaByID(objSelProdotto.getIDTassaApplicata()).getDescrizioneTassa()
              end if
            end if
						
            if(descTassa<>"") then descTassa = "&nbsp;&nbsp;("&descTassa&")" end if   
           
            
						'************ converto il prezzo in base alla valuta selezionata
						if not(defCurrObj="") AND not(thisCurrObj="") then
							numPrezzoOld = currClass.convertCurrency(numPrezzoOld, defCurrObj, thisCurrObj)
							numPrezzoReal = currClass.convertCurrency(numPrezzoReal, defCurrObj, thisCurrObj)
						end if
						%>
                          
						<%if(discountPercent > 0) then%><span class="testoBarrato"><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(numPrezzoOld, 2,-1)&descTassa%></span> --><%end if%>&nbsp;<%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(numPrezzoReal, 2,-1)&descTassa%>
						<%if(discountPercent > 0) then response.Write("<br><br>" & lang.getTranslated("frontend.template_prodotto.table.label.sconto") & " " & discountPercent & "%") end if%>
						<%if(objSelProdotto.getAttivo() = 0) then response.Write("<br><br>" & lang.getTranslated("frontend.template_prodotto.table.alt.non_disponibile")) end if%>
						
						<%
						On Error Resume Next
						hasProdFields=false
						
						Set objListProdField = objProdField.getListProductField4ProdActive(objSelProdotto.getIDProdotto())
						
						if(objListProdField.Count > 0)then
							hasProdFields = true
						end if
						
						if(Err.number <> 0) then
							hasProdFields = false
						end if	
						
						if(hasProdFields)then
							
							tmpGroupDesc = ""	
							tmpGroupDescTrans = ""	
							
							response.write("<p>")
										
							for each k in objListProdField
								On Error Resume next
								Set objField = objListProdField(k)
							
								labelForm = objField.getDescription()
								if not(lang.getTranslated("frontend.prodotto.field.label."&objField.getDescription())="") then labelForm = lang.getTranslated("frontend.prodotto.field.label."&objField.getDescription())
								
								'*** imposto la descrizione per il gruppo di appartenenza
								if(strComp(typename(objField.getObjGroup()), "ProductFieldGroupClass") = 0)then
									tmpDescG = objField.getObjGroup().getDescription()
									if(tmpDescG <> tmpGroupDesc)then
										tmpGroupDesc = tmpDescG
                    					tmpGroupDescTrans = tmpGroupDesc
										if not(lang.getTranslated("frontend.prodotto.field.label.group."&tmpGroupDesc)="") then tmpGroupDescTrans = lang.getTranslated("frontend.prodotto.field.label.group."&tmpGroupDesc)
										
										labelForm = "<div class=""prodotto_field_prod_group"">"& tmpGroupDescTrans & "</div>" & labelForm
									end if
								end if
									
								fieldCssClass=""
	
								select Case objField.getTypeField()								
								Case 1,2
									fieldCssClass="formFieldTXTMedium"
									if(objField.getEditable()="1")then
										response.write(labelForm & ":&nbsp;" &objProdField.renderProductFieldHTML(objField,fieldCssClass, ProdottoCounter, objSelProdotto.getIDProdotto(), "",lang,1,objField.getEditable()) & "<br/>")%>
									<%
									else
										valueTmp = objField.getSelValue()
										response.write(labelForm & ":&nbsp;" & valueTmp & "<br/>")
									end if
								Case 3,4,5,6							
									if(CInt(objField.getTypeField())=4) then
										fieldCssClass="formFieldMultiple"
									end if
									response.write(labelForm & ":&nbsp;" &objProdField.renderProductFieldHTML(objField,fieldCssClass, ProdottoCounter, objSelProdotto.getIDProdotto(), "",lang,1,objField.getEditable()) & "<br/>")
                Case 7
                  fieldValueMatch = objProdField.findFieldMatchValue(k,objSelProdotto.getIDProdotto())
                  response.write(objProdField.renderProductFieldHTML(objField,fieldCssClass, ProdottoCounter, objSelProdotto.getIDProdotto(), fieldValueMatch,lang,1,objField.getEditable()))
								Case 8									
									fieldCssClass="formFieldTXTMedium"
									if(objField.getEditable()="1")then
										response.write(labelForm & ":&nbsp;" &objProdField.renderProductFieldHTML(objField,fieldCssClass, ProdottoCounter, objSelProdotto.getIDProdotto(), "",lang,1,objField.getEditable()) & "<br/>")%>
									<%
									else
										valueTmp = objField.getSelValue()
										response.write(labelForm & ":&nbsp;" & valueTmp & "<br/>")
									end if
                Case 9
									fieldCssClass="formFieldTXTMedium"
									if(objField.getEditable()="1")then%>
                  <script>
                  //declare cleditor option array;
                  var cloptions<%=objProdField.getFieldPrefix()&ProdottoCounter&objField.getID()%> = {
                  width:280,	// width not including margins, borders or padding
                  height:200,	// height not including margins, borders or padding
                  controls:"bold italic underline strikethrough subscript superscript | font size style | color highlight removeformat | bullets numbering | alignleft center alignright justify | rule | cut copy paste | image",	// controls to add to the toolbar
                  }
                  </script>	
										<%response.write(labelForm & ":&nbsp;" &objProdField.renderProductFieldHTML(objField,fieldCssClass, ProdottoCounter, objSelProdotto.getIDProdotto(), "",lang,1,objField.getEditable()) & "<br/>")%>
									<%
									else
										valueTmp = Server.HTMLEncode(objField.getSelValue())
										response.write(labelForm & ":&nbsp;" & valueTmp & "<br/>")
									end if                  
                Case Else
								End Select
			
								Set objField = nothing
	
								if(Err.number<>0) then
								'response.write(Err.description)
								end if
							next
							
							response.write("</p>")						
							
							Set objListProdField = nothing
						end if				
						%>					
						
						</p>
            <%end if%>
					</div>
					<div id="clear"></div>
					<div id="prodotto-footer">
						<%if(Application("disable_ecommerce") = 0) then%>
              <%'GESTISCO LA QUANTITA' SELEZIONABILE								
              if(objSelProdotto.getQtaDisp() = Application("unlimited_key")) then%>
                <input type="text" name="qta_prodotto" value="" onkeypress="javascript:return isInteger(event);" class="formFieldTXTShort">
              <%else%>
                <!--<select name="qta_prodotto" class="formFieldSelect">
                <option value=""></option>
                <%'for statiOrderCount = 1 to objSelProdotto.getQtaDisp()%>
                  <option value="<%'=statiOrderCount%>" ><%'=statiOrderCount%></option>
                <%'next%>
                </select>-->
                <input type="text" name="qta_prodotto" id="qta_prodotto" value="" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);" onblur="javascript:checkMaxQtaProd(<%=objSelProdotto.getQtaDisp()%>,this);">&nbsp;<%=lang.getTranslated("frontend.template_prodotto.table.label.product_disp")&"&nbsp;"&objSelProdotto.getQtaDisp()%>&nbsp;&nbsp;		
              <%end if%>
                
              <%if(objSelProdotto.getAttivo() = 1) then%>
              <a href="javascript:addToCarrello(document.form_add_to_carrello_<%=ProdottoCounter%>,<%=ProdottoCounter%>);" title="<%=lang.getTranslated("frontend.template_prodotto.table.alt.put_on_carrello")%>"><img src="<%=Application("baseroot")&"/common/img/cart_add.png"%>" width="16" height="16" hspace="0" vspace="0" border="0"><%=lang.getTranslated("frontend.template_prodotto.table.alt.put_on_carrello")%></a>
              <%else%>
              <img src="<%=Application("baseroot")&"/common/img/ico-carrello.png"%>" hspace="0" vspace="0" border="0" width="16" height="16" alt="<%=lang.getTranslated("frontend.template_prodotto.table.alt.non_disponibile")%>"><br><br>
              <%end if%>	
            <%end if%>
						<%if(not(isNull(objCommento.findCommentiByIDElement(objSelProdotto.getIDProdotto(),2,1)))) AND (Instr(1, typename(objCommento.findCommentiByIDElement(objSelProdotto.getIDProdotto(),2,1)), "dictionary", 1) > 0) then%>
						&nbsp;|&nbsp;<a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popupComments.asp?id_element="&objSelProdotto.getIDProdotto()&"&element_type=2&active=1"%>','popupallegati',400,400,100,100);" title="<%=lang.getTranslated("frontend.template_prodotto.table.alt.see_comments")%>"><img src="<%=Application("baseroot")&"/common/img/ico-commenti.png"%>" hspace="0" vspace="0" width="16" height="16" border="0"><%=lang.getTranslated("frontend.template_prodotto.label.see_comments")%></a>
						<%end if%>
					</div>		
					</form>
					<%Set objSelProdotto = nothing%>
					</div>
				<%next				
				Set objTasse = nothing%>
				<div>
				<%if(totPages > 1) then%>
					<%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("controller_page"), "&order_by="&order_prodotto_by&"&modelPageNum="&modelPageNum)
				end if%>
				</div>
			</div>
		<%else
			response.Write("<br/><br/><div align=""center""><strong>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</strong></div>")
		end if%>		