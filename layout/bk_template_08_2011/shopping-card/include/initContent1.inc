		<%if(not(request("error")) = "" AND request("error") = 1) then%>
				<br><span class="labelTabCarrello"><%=lang.getTranslated("frontend.carrello.table.label.error_wrong_qta") & " " & request("nome_prod")%></span><br><br>
		<%end if
		
		'************** codice per la lista Prodotto e paginazione
		if(not(isNull(objCarrelloUser))) then
			id_carrello = objCarrelloUser.getIDCarrello()
			Dim objProdPerCarrello
			Set objProdPerCarrello = New ProductsCardClass
			
			bolHasObj = false
			
			on error Resume Next
			Set objListaCarrello = objProdPerCarrello.retrieveListaProdotti(objCarrelloUser.getIDCarrello())	
			
			if(objListaCarrello.Count > 0) then		
				bolHasObj = true
			end if
				
			if Err.number <> 0 then
				bolHasObj = false
			end if	
			
			Dim totaleProdottoImp4spese, totale_carrello
			totaleProdottoImp4spese = 0
			totale_carrello = 0
			
			if(bolHasObj) then
				Dim objTmpProdComplete, objSelProdotto, objListaFile, ProdottoCounter, iIndex 
				Dim objTmpProdotto, FromProdotto, ToProdotto, Diff
				Dim totaleProdottoImp, totaleProdottoTax, marginPercent, discountPercent, totMarginAmount, totDiscountAmount, singleMarginAmount
				objYmpKey = objListaCarrello.Keys%>	
				
				<div id="carrello-lista"><%
					Dim counter_card ,descTassa
					counter_card = 0
          totMarginAmount = 0
          totDiscountAmount = 0
					
					for each k in objYmpKey
						Set objSelProdotto = objListaCarrello.item(k)
						Set objTmpProdComplete = Prodotto.findProdottoByID(objSelProdotto.getIDProd(),1)
						totaleProdottoImp = 0
						totaleProdottoTax = 0
            marginPercent = 0
            discountPercent = 0
            singleMarginAmount = 0

						if(hasGroup) then
						  On Error Resume Next
						  Dim objSelMargin
			
						  Set objSelMargin = objGroup.getMarginDiscountXUserGroup(groupCliente)
              marginPercent = objSelMargin.getMargin()
              discountPercent = objSelMargin.getDiscountPercentual(CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objTmpProdComplete.getsconto()),CDbl(scontoCliente))
              totaleProdottoImp = CDbl(objTmpProdComplete.getPrezzo()) * objSelProdotto.getQtaProd()
              singleMarginAmount = objSelMargin.getMarginAmount(totaleProdottoImp,CDbl(objSelMargin.getMargin()))
						  totMarginAmount = totMarginAmount + singleMarginAmount
              totDiscountAmount = totDiscountAmount + objSelMargin.getDiscountAmount(totaleProdottoImp,CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objTmpProdComplete.getsconto()),CDbl(scontoCliente))
						  totaleProdottoImp = objSelMargin.getAmount(totaleProdottoImp,CDbl(objSelMargin.getMargin()),CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objTmpProdComplete.getsconto()),CDbl(scontoCliente))
						  if(Err.number <>0) then
						  end if	
						  Set objSelMargin = nothing
						else          
						  if(objTmpProdComplete.hasSconto() AND (not(hasSconto) OR (hasSconto AND Application("manage_sconti") = 1))) then
							totaleProdottoImp = CDbl(objTmpProdComplete.getPrezzoScontato()) * objSelProdotto.getQtaProd()
              discountPercent = CDbl(objTmpProdComplete.getsconto())
							if(hasSconto)then
							  totaleProdottoImp = totaleProdottoImp - (totaleProdottoImp / 100 * scontoCliente)	
                discountPercent = discountPercent+CDbl(scontoCliente)						
							end if
						  else
							totaleProdottoImp = CDbl(objTmpProdComplete.getPrezzo()) * objSelProdotto.getQtaProd()
							if(hasSconto)then
							  totaleProdottoImp = totaleProdottoImp - (totaleProdottoImp / 100 * scontoCliente)				
                discountPercent = CDbl(scontoCliente)			
							end if
						  end if
						end if


            '***********************************   INTERNAZIONALIZZAZIONE TASSE   ****************************
						descTassa = ""
            applyOrigTax = true
            if(Application("enable_international_tax_option")=1) AND (international_country_code<>"") then
              if(hasGroup AND (Instr(1, typename(groupClienteTax), "TaxsGroupClass", 1) > 0)) then
				On Error Resume Next
				objRelatedTax = groupClienteTax.findRelatedTax(groupClienteTax.getID(), international_country_code,international_state_region_code)

                if(not(isNull(objRelatedTax))) then
                  Set objTaxG = objTasse.findTassaByID(objRelatedTax)
                  totaleProdottoTax = groupClienteTax.getImportoTassa(totaleProdottoImp, objTaxG)
                  descTassa = objTaxG.getDescrizioneTassa()
                  Set objTaxG = nothing
                  applyOrigTax = false
                else
                  applyOrigTax = true		
                end if		
				if(Err.number<>0)then
					applyOrigTax = true
				end if
              else
				On Error Resume Next
				Set groupProdTax = objTmpProdComplete.getTaxGroupObj(objTmpProdComplete.getTaxGroup()) 
                if(Instr(1, typename(groupProdTax), "TaxsGroupClass", 1) > 0) then
				  objRelatedTax = groupProdTax.findRelatedTax(groupProdTax.getID(), international_country_code,international_state_region_code)				  
				  if(not(isNull(objRelatedTax))) then
					Set objTaxG = objTasse.findTassaByID(objRelatedTax)
                    totaleProdottoTax = groupProdTax.getImportoTassa(totaleProdottoImp, objTaxG)
                    descTassa = objTaxG.getDescrizioneTassa()
                    Set objTaxG = nothing
                    applyOrigTax = false		
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
              totaleProdottoTax = 0
              taxDesc = ""
              if not(isNull(objTmpProdComplete.getIDTassaApplicata())) AND not(objTmpProdComplete.getIDTassaApplicata() = "") then
                totaleProdottoTax = objTmpProdComplete.getImportoTassa(totaleProdottoImp)
                descTassa = objTasse.findTassaByID(objTmpProdComplete.getIDTassaApplicata()).getDescrizioneTassa()
              end if
            end if
						
            if(descTassa<>"") then descTassa = "&nbsp;&nbsp;("&descTassa&")" end if
            
            
						totaleProdottoImp4spese = totaleProdottoImp4spese+totaleProdottoImp
						totale_carrello = totale_carrello+totaleProdottoImp+totaleProdottoTax
						%>
						<div>
						<form action="<%=Application("baseroot")&Application("dir_upload_templ")&"shopping-card/ManageCarrello.asp"%>" method="post" name="form_carrello_<%=counter_card%>" enctype="multipart/form-data">	
						<input type="hidden" value="<%=objSelProdotto.getIDProd()%>" name="id_prodotto">
						<input type="hidden" value="<%=objSelProdotto.getCounterProd()%>" name="counter_prod">		
						<input type="hidden" value="del" name="operation">
						<input type="hidden" value="<%=objSelProdotto.getQtaProd()%>" name="qta_prodotto">			
						<div id="prodotto-immagine">
							<%if (not(isNull(objTmpProdComplete.getFileXProdotto())) AND not(isEmpty(objTmpProdComplete.getFileXProdotto()))) then
								Dim hasNotSmallImg
								hasNotSmallImg = true
								Set objListaFilePerProdotto = objTmpProdComplete.getFileXProdotto()					
								for each xObjFile in objListaFilePerProdotto
									Set objFileXProdotto = objListaFilePerProdotto(xObjFile)
									iTypeFile = objFileXProdotto.getFileTypeLabel()
									if(Cint(iTypeFile) = 1) then%>	
										<img src="<%=Application("dir_upload_prod")&objFileXProdotto.getFilePath()%>" hspace="0" vspace="0" border="0" width="50" height="50" alt="<%=objTmpProdComplete.getNomeProdotto()%>" />
										<%hasNotSmallImg = false
										Exit for
									end if
									Set objFileXProdotto = nothing	
								next
								
								if(hasNotSmallImg) then%>
								<img width="50" height="50" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0">
								<%end if
								Set objListaFilePerProdotto = nothing		
							else%>
							<img width="50" height="50" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0">
							<%end if%>
						</div>
						<div id="prodotto-carrello">
							<h2><%=objTmpProdComplete.getNomeProdotto()%></h2>
							
							<%
							if (Instr(1, typename(objProdField.findListFieldXCardByProd(objSelProdotto.getCounterProd(), id_carrello, objSelProdotto.getIDProd())), "Dictionary", 1) > 0) then
								Set fieldList4Card = objProdField.findListFieldXCardByProd(objSelProdotto.getCounterProd(), id_carrello, objSelProdotto.getIDProd())			
							
								if(fieldList4Card.count > 0)then%>
									<p>											
									<%for each q in fieldList4Card
										
										Set objTmpField4Card = fieldList4Card(q)
										keys = objTmpField4Card.Keys
										
										for each r in keys
											Set tmpF4O = r
																				
											labelTmp = ""
											labelTmp = tmpF4O.getDescription()
											if(Cint(tmpF4O.getTypeField())<>9)then
												valueTmp = Server.HTMLEncode(tmpF4O.getSelValue())
											else
												valueTmp = tmpF4O.getSelValue()
											end if
											if(Cint(tmpF4O.getTypeField())=8)then
												valueTmp = "<a href=""" & valueTmp & """ target=_blank>click</a>"
											end if
											if not(lang.getTranslated("frontend.prodotto.field.label."&tmpF4O.getDescription())="") then labelTmp = lang.getTranslated("frontend.prodotto.field.label."&tmpF4O.getDescription())
											response.write(labelTmp & ":&nbsp;" & valueTmp & "<br/>")											
																		
											Set tmpF4O = nothing
										next
										Set objTmpField4Card = nothing							
									next%>
									</p>
								<%end if
								Set fieldList4Card = nothing
							end if
							%>							
							
							<p><strong><%=lang.getTranslated("frontend.carrello.table.label.quantita")%>: </strong><%=objSelProdotto.getQtaProd()%></p>
							<%
							'************ converto il prezzo in base alla valuta selezionata
							if(hasCurrency) then
								numPrezzoProd = currClass.convertCurrency(totaleProdottoImp+totaleProdottoTax, defCurrObj, Session("currency"))
                singleMarginAmount = currClass.convertCurrency(singleMarginAmount, defCurrObj, Session("currency"))
							else
								numPrezzoProd = totaleProdottoImp+totaleProdottoTax
							end if
							%>
							<p><strong><%=lang.getTranslated("frontend.carrello.table.label.totale")%>: </strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(numPrezzoProd, 2,-1)&descTassa%>
              &nbsp;<a href="javascript:showHideDiv('prod-commissions-<%=counter_card%>');">?</a>
              <div id="prod-commissions-<%=counter_card%>" style="<%if ((hasGroup AND singleMarginAmount > 0) OR (discountPercent > 0)) then%>margin-bottom:3px;padding:10px;vertical-align:middle;text-align:left;font-size: 10px;text-decoration: none;border:0px solid;background:#FFFFFF;width:160px;left:-10px;top:0px;<%end if%>visibility:hidden;display:none;position:relative;">
                <%if (hasGroup AND singleMarginAmount > 0) then%>
                <%=lang.getTranslated("frontend.carrello.table.label.commissioni")%>: <%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(singleMarginAmount, 2,-1)%><%'="&nbsp;"&marginPercent&"%"%><br/>
                <%end if%>
                <%if (discountPercent > 0) then%>
                <%=lang.getTranslated("frontend.carrello.table.label.sconto_applicato")%>: <%=discountPercent&"%"%>
                <%end if%>
              </div>
              </p>
						</div>
						<div id="prodotto-cancella"><a href="javascript:delFromCarrello(document.form_carrello_<%=counter_card%>);"><span><%=lang.getTranslated("frontend.carrello.table.label.del_prod")%></span></a></div>
						<div id="clear"></div>
						<div id="prodotto-footer"></div>
						</form>									
						<%Set objTmpProdComplete = nothing
						Set objSelProdotto = nothing
						counter_card = counter_card +1%>
						</div>
					<%next%>	
					
					<div id="prodotto-totale">
						<%
						'****************  RIVEDERE CALCOLI IMPONIBILE E TASSE PER SINGOLI PRODOTTI
						'****************  SPESE SPEDIZIONE E TOTALE
						totale_carrello = FormatNumber(totale_carrello, 2,-1)%>
						<%   
            if not(hasGroup) then           
						  if(hasSconto) then
							response.write("<br/>"&lang.getTranslated("frontend.carrello.table.label.sconto_cliente") & ": " &  scontoCliente&" %")
						  end if
											  
							if(hasSconto AND Application("manage_sconti") = 0) then
							  response.write("<br/>"&lang.getTranslated("frontend.carrello.table.label.if_client_has_sconto")&"<br/>")
							end if
						end if%>
            <br/>
						<form action="<%=Application("baseroot") &Application("dir_upload_templ")&"shopping-card/carrello2.asp"%>" method="post" name="form_insert_carrello">
						<input type="hidden" value="<%=id_carrello%>" name="id_carrello">
						<%
						'************ converto il prezzo in base alla valuta selezionata
						if(hasCurrency) then
              totMarginAmount = currClass.convertCurrency(totMarginAmount, defCurrObj, Session("currency"))
              totDiscountAmount = currClass.convertCurrency(totDiscountAmount, defCurrObj, Session("currency"))
							totale_carrello = currClass.convertCurrency(totale_carrello, defCurrObj, Session("currency"))
						end if
						%>
            <%if (hasGroup) then%>
            <strong><%=lang.getTranslated("frontend.carrello.table.label.totale_commissioni")%>:&nbsp;</strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(totMarginAmount, 2,-1)%><br/>
            <strong><%=lang.getTranslated("frontend.carrello.table.label.totale_sconti")%>:&nbsp;</strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(totDiscountAmount, 2,-1)%><br/>
            <br/>
            <%end if%>
						<strong><%=lang.getTranslated("frontend.carrello.table.label.totale_prodotti")%>:&nbsp;</strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(totale_carrello, 2,-1)%>&nbsp;
						<a href="javascript:sendCarrello();"><%=lang.getTranslated("frontend.carrello.table.label.continue_order")%></a>&nbsp;&nbsp;
						<a href="<%=Application("baseroot") &Application("dir_upload_templ")&"shopping-card/DeleteCarrello.asp?id_carrello_to_delete=" & id_carrello%>"><%=lang.getTranslated("frontend.carrello.table.label.cancel_card")%></a>
						</form>	
					</div>
				</div>
				<%Set objListaCarrello = nothing			
			else
				response.write("<br><br><p align='center'>"&lang.getTranslated("frontend.carrello.table.label.empty_card")&"</p>")
			end if
		else			
			response.write("<br><br><p align='center'>"&lang.getTranslated("frontend.carrello.table.label.empty_card")&"</p>")
		end if
    
    Set objGroup = nothing%>