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
				Dim totaleProdottoImp, totaleProdottoTax, marginPercent, discountPercent
				Dim totMarginAmount, totDiscountAmount, singleMarginAmount
				objYmpKey = objListaCarrello.Keys%>	
				
				<div id="carrello-lista"><%
					Dim counter_card ,descTassa, applyBills
					counter_card = 0
					totMarginAmount = 0
					totDiscountAmount = 0
					
					applyBills = false
      
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
                marginPercent = objSelMargin.getMargin()
                discountPercent = objSelMargin.getDiscountPercentual(CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objTmpProdComplete.getsconto()),CDbl(scontoCliente))
                totaleProdottoImp = CDbl(objTmpProdComplete.getPrezzo()) * objSelProdotto.getQtaProd()
                singleMarginAmount = objSelMargin.getMarginAmount(totaleProdottoImp,CDbl(objSelMargin.getMargin()))
                totMarginAmount = totMarginAmount + singleMarginAmount
                totDiscountAmount = totDiscountAmount + objSelMargin.getDiscountAmount(totaleProdottoImp,CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objTmpProdComplete.getsconto()),CDbl(scontoCliente))
                totaleProdottoImp = objSelMargin.getAmount(totaleProdottoImp,CDbl(objSelMargin.getMargin()),CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objTmpProdComplete.getsconto()),CDbl(scontoCliente))
                if(Err.number <>0) then
                end if	
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
                ' verifico se l'utente ha selezionato il flag tipologia cliente=societ√† e se per il country/region selezionato il falg escludi tassa √® attivo
                if(Cint(userIsCompanyClient)=1 AND groupClienteTax.isTaxExclusion(groupClienteTax.getID(), international_country_code,international_state_region_code))then
                  totaleProdottoTax = 0
                  descTassa = lang.getTranslated("frontend.prodotti.label.tax_excluded")							
                  applyOrigTax = false
                else
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
                end if
                if(Err.number<>0)then
                  applyOrigTax = true
                end if
              else
				On Error Resume Next
				Set groupProdTax = objTmpProdComplete.getTaxGroupObj(objTmpProdComplete.getTaxGroup()) 
				if(Instr(1, typename(groupProdTax), "TaxsGroupClass", 1) > 0) then
          ' verifico se l'utente ha selezionato il flag tipologia cliente=societ√† e se per il country/region selezionato il falg escludi tassa √® attivo
          if(Cint(userIsCompanyClient)=1 AND groupProdTax.isTaxExclusion(groupProdTax.getID(), international_country_code,international_state_region_code))then
            totaleProdottoTax = 0
            descTassa = lang.getTranslated("frontend.prodotti.label.tax_excluded")							
            applyOrigTax = false
          else	
            objRelatedTax = groupProdTax.findRelatedTax(groupProdTax.getID(), international_country_code,international_state_region_code)				  
            if(not(isNull(objRelatedTax))) then
            Set objTaxG = objTasse.findTassaByID(objRelatedTax)
            totaleProdottoTax = groupProdTax.getImportoTassa(totaleProdottoImp, objTaxG)
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
              totaleProdottoTax = 0
              taxDesc = ""
              if not(isNull(objTmpProdComplete.getIDTassaApplicata())) AND not(objTmpProdComplete.getIDTassaApplicata() = "") then
                totaleProdottoTax = objTmpProdComplete.getImportoTassa(totaleProdottoImp)
                descTassa = objTasse.findTassaByID(objTmpProdComplete.getIDTassaApplicata()).getDescrizioneTassa()
              end if
            end if
						
            if(descTassa<>"") then descTassa = "&nbsp;&nbsp;("&descTassa&")" end if             
				
					'************ se il prodotto non Ë di tipo scaricabile aggiorno l'imponibile su cui verranno calcolate le spese di spedizione
					if not(objTmpProdComplete.isDownloadable()=1) then
					  totaleProdottoImp4spese = totaleProdottoImp4spese+totaleProdottoImp
					  applyBills = true
					end if
            
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
					
					<div id="prodotto-conto">
						<%
						'****************  RIVEDERE CALCOLI IMPONIBILE E TASSE PER SINGOLI PRODOTTI
						'****************  SPESE SPEDIZIONE E TOTALE
						totale_carrello = FormatNumber(totale_carrello, 2,-1)
							
						'************ converto il prezzo in base alla valuta selezionata
						if(hasCurrency) then
              totMarginAmount = currClass.convertCurrency(totMarginAmount, defCurrObj, Session("currency"))
              totDiscountAmount = currClass.convertCurrency(totDiscountAmount, defCurrObj, Session("currency"))
							totale_carrello_tmp = currClass.convertCurrency(totale_carrello, defCurrObj, Session("currency"))
						else
							totale_carrello_tmp = totale_carrello
						end if%>
						<div class="spese-div">
            <%if (hasGroup) then%>
            <%=lang.getTranslated("frontend.carrello.table.label.totale_commissioni")%>:&nbsp;<strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(totMarginAmount, 2,-1)%></strong><br/>
            <%=lang.getTranslated("frontend.carrello.table.label.totale_sconti")%>:&nbsp;<strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(totDiscountAmount, 2,-1)%></strong><br/>
            <br/>
            <%end if%>
            <%=lang.getTranslated("frontend.carrello.table.label.totale_prodotti")%>:&nbsp;<strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(totale_carrello_tmp, 2,-1)%></strong>
						<br/>
            <%
            if not(hasGroup) then
              if(hasSconto) then
                response.write("<br/>"&lang.getTranslated("frontend.carrello.table.label.sconto_cliente") & ": " &  scontoCliente&" %")
              end if
                
                if(hasSconto AND Application("manage_sconti") = 0) then
                  response.write("<br/>"&lang.getTranslated("frontend.carrello.table.label.if_client_has_sconto")&"<br/><br/>")
                end if
            end if%>              
            </div>


						<form action="<%=Application("baseroot") &Application("dir_upload_templ")&"shopping-card/ProcessCarrello.asp"%>" method="post" name="form_insert_carrello">
                
            <%
						On Error Resume Next
						Dim objBillsClass, totSpeseImp, totSpeseTax, totSpese
						Dim objListaSpeseXCarrello, objTmpSpesa, objTmpSpesaXCarrello
						Set objBillsClass = new BillsClass		
						Set objListaSpeseXCarrello = objBillsClass.getListaSpese(null, null, 1, null)
						totSpese = 0
						totSpeseAuto = 0
						
            if(applyBills) then%>
						<div class="spese-div">
            <strong><%=lang.getTranslated("frontend.carrello.table.label.spese_spedizione")%>:</strong><br/> 
              <%
              oldGroupDesc = ""
            
              for each j in objListaSpeseXCarrello.Keys
                Set objTmpSpesaXCarrello = objListaSpeseXCarrello(j)
                totSpeseImp = 0
                totSpeseTax = 0
              
                if(oldGroupDesc<>objTmpSpesaXCarrello.getGroup())then 
                  if not(lang.getTranslated("portal.commons.order_bills.label.group."&objTmpSpesaXCarrello.getGroup())="") then response.write("<b>"&lang.getTranslated("portal.commons.order_bills.label.group."&objTmpSpesaXCarrello.getGroup())&"</b><br/>") else response.write("<b>"&objTmpSpesaXCarrello.getGroup()&"</b><br/>") end if	
                end if
              
                if(CInt(objTmpSpesaXCarrello.getTipoValore()) = 2) then
                  totSpeseImp = CDbl(totaleProdottoImp4spese) / 100 * CDbl(objTmpSpesaXCarrello.getValore())
                else
                  totSpeseImp = CDbl(objTmpSpesaXCarrello.getValore())
                end if                  


                '***********************************   INTERNAZIONALIZZAZIONE TASSE   ****************************
                applyOrigTax = true
                if(Application("enable_international_tax_option")=1) AND (international_country_code<>"") then
                  if(hasGroup AND (Instr(1, typename(groupClienteTax), "TaxsGroupClass", 1) > 0)) then
					On Error Resume Next
					objRelatedTax = groupClienteTax.findRelatedTax(groupClienteTax.getID(), international_country_code,international_state_region_code)
                    if(not(isNull(objRelatedTax))) then
                      Set objTaxG = objTasse.findTassaByID(objRelatedTax)
                      totSpeseTax = groupClienteTax.getImportoTassa(totSpeseImp, objTaxG)
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
                    Set groupBillsTax = objTmpSpesaXCarrello.getTaxGroupObj(objTmpSpesaXCarrello.getTaxGroup())
                    if(Instr(1, typename(groupBillsTax), "TaxsGroupClass", 1) > 0) then
					  objRelatedTax = groupBillsTax.findRelatedTax(groupBillsTax.getID(), international_country_code,international_state_region_code)
                      if(not(isNull(objRelatedTax))) then
                        Set objTaxG = objTasse.findTassaByID(objRelatedTax)
                        totSpeseTax = groupBillsTax.getImportoTassa(totSpeseImp, objTaxG)
                        Set objTaxG = nothing
                        applyOrigTax = false		
                      end if								
                    else
                      applyOrigTax = true
                    end if
                    Set groupBillsTax = nothing	
					if(Err.number<>0)then
                      applyOrigTax = true
					end if	
                  end if
                end if
                if(applyOrigTax)then
                  totSpeseTax = 0
                  if not(isNull(objTmpSpesaXCarrello.getIDTassaApplicata())) AND not(objTmpSpesaXCarrello.getIDTassaApplicata() = "") then
                    Set objBillTaxTmp = objTasse.findTassaByID(objTmpSpesaXCarrello.getIDTassaApplicata())
                    if(objBillTaxTmp.getTipoValore() = 2) then
                      totSpeseTax = CDbl(totSpeseImp) * (CDbl(objBillTaxTmp.getValore()) / 100)
                    else
                      totSpeseTax = CDbl(objBillTaxTmp.getValore())
                    end if	
                    Set objBillTaxTmp = nothing
                  end if
                end if	                  
                  

                '************ converto il prezzo in base alla valuta selezionata
                if(hasCurrency) then
                  totale_spese_tmp = currClass.convertCurrency(totSpeseImp+totSpeseTax, defCurrObj, Session("currency"))
                else
                  totale_spese_tmp = totSpeseImp+totSpeseTax
                end if
                
                 
                if(objTmpSpesaXCarrello.getAutoactive()=1)then
                  totSpese = totSpese+totSpeseImp+totSpeseTax
                  totSpeseAuto = totSpeseAuto+totSpeseImp+totSpeseTax                  

                  if not(lang.getTranslated("portal.commons.order_bills.label."&objTmpSpesaXCarrello.getDescrizioneSpesa())="") then response.write("&nbsp;"&lang.getTranslated("portal.commons.order_bills.label."&objTmpSpesaXCarrello.getDescrizioneSpesa())&"&nbsp;&nbsp;&nbsp;<strong>") else response.write("&nbsp;"&objTmpSpesaXCarrello.getDescrizioneSpesa()&"&nbsp;&nbsp;&nbsp;<strong>") end if
                  if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then 
                    response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) 
                  else 
                    response.write(Session("currency")) 
                  end if
                  response.write("&nbsp;"&FormatNumber(totale_spese_tmp,2,-1,-2,0)&"</strong><br>")                  
                else%>	
                  <script language="Javascript">
                  listBills4Order.put("<%=objTmpSpesaXCarrello.getGroup()&"-"&j&"-"&objTmpSpesaXCarrello.getRequired()%>","<%=CDbl(totSpeseImp)+CDbl(totSpeseTax)%>");
                  </script>
                  <%if(objTmpSpesaXCarrello.getMultiply()=1)then%>
                    <input type="checkbox" onclick="javascript:calculateBills4Order('<%=Cdbl(totale_carrello)+Cdbl(totSpeseAuto)%>','<%=currClass.findCurrencyByCurrency(defCurrObj).getRate()%>','<%if(Trim(Session("currency")) <> "") then response.write(currClass.findCurrencyByCurrency(Session("currency")).getRate()) else response.write(currClass.findCurrencyByCurrency(defCurrObj).getRate()) end if%>');" name="<%=objTmpSpesaXCarrello.getGroup()%>" id="<%=objTmpSpesaXCarrello.getGroup()&"-"&j&"-"&objTmpSpesaXCarrello.getRequired()%>" value="<%=j%>"/> <%if not(lang.getTranslated("portal.commons.order_bills.label."&objTmpSpesaXCarrello.getDescrizioneSpesa())="") then response.write(lang.getTranslated("portal.commons.order_bills.label."&objTmpSpesaXCarrello.getDescrizioneSpesa())) else response.write(objTmpSpesaXCarrello.getDescrizioneSpesa()) end if%>&nbsp;&nbsp;&nbsp;<%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(totale_spese_tmp,2,-1)%>&nbsp;&nbsp;<BR>
                  <%else%>
                    <input type="radio"  onclick="javascript:calculateBills4Order('<%=Cdbl(totale_carrello)+Cdbl(totSpeseAuto)%>','<%=currClass.findCurrencyByCurrency(defCurrObj).getRate()%>','<%if(Trim(Session("currency")) <> "") then response.write(currClass.findCurrencyByCurrency(Session("currency")).getRate()) else response.write(currClass.findCurrencyByCurrency(defCurrObj).getRate()) end if%>');" name="<%=objTmpSpesaXCarrello.getGroup()%>" id="<%=objTmpSpesaXCarrello.getGroup()&"-"&j&"-"&objTmpSpesaXCarrello.getRequired()%>" value="<%=j%>"/> <%if not(lang.getTranslated("portal.commons.order_bills.label."&objTmpSpesaXCarrello.getDescrizioneSpesa())="") then response.write(lang.getTranslated("portal.commons.order_bills.label."&objTmpSpesaXCarrello.getDescrizioneSpesa())) else response.write(objTmpSpesaXCarrello.getDescrizioneSpesa()) end if%>&nbsp;&nbsp;&nbsp;<%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(totale_spese_tmp,2,-1)%>&nbsp;&nbsp;<BR>					
                  <%end if
                  
                end if
                  
                oldGroupDesc = objTmpSpesaXCarrello.getGroup()
                Set objTmpSpesaXCarrello = nothing
              next%>
            </div>
            <%end if
						
						Set objListaSpeseXCarrello = nothing
						Set objBillsClass = nothing
						
						If Err.Number<>0 then
						end if
						%>
						
            <div class="spese-div">
						<input type="hidden" value="<%=id_carrello%>" name="id_carrello">
						  <strong><%=lang.getTranslated("frontend.carrello.table.label.tipo_pagam_order")%>:</strong><br>	
						  <div id="payment_list">
						  <%
               On Error Resume Next
						   Dim objPayment, objTmpPayment, objListaPayment, objModuloPayment, objModulo, strLogo, loadDirectPayment		   
               loadDirectPayment = null
               if((totSpese+totale_carrello)<=0) then
                loadDirectPayment = 0
               end if
						   Set objPayment = New PaymentClass
						   Set objModulo = new PaymentModuleClass					   
						   Set objListaPayment = objPayment.getListaPayment(1,loadDirectPayment)
              if not(isEmpty(objListaPayment)) AND (Instr(1, typename(objListaPayment), "Dictionary", 1) > 0) then%>

						<script language="Javascript">
						<%for each k in objListaPayment.Keys%>
						listPaymentMethods.put("<%=k%>","<%=objListaPayment(k).getCommission()&"|"&objListaPayment(k).getCommissionType()%>");	
						<%next%>
						</script>
						<ul>
						  <%for each k in objListaPayment.Keys
							strLogo = ""
							if(Cint(objListaPayment(k).getPaymentModuleID()) <> -1) then
								Set objModuloPayment = objModulo.findPaymentModuloByID(objListaPayment(k).getPaymentModuleID())
								strLogo = objModuloPayment.getLogo()
								Set objModuloPayment = nothing
							end if%>
								<li><INPUT type="radio" name="tipo_pagam" value="<%=k%>" onclick="javascript:calculatePaymentCommission('<%=Cdbl(totale_carrello)+Cdbl(totSpeseAuto)%>',<%=k%>,'<%=currClass.findCurrencyByCurrency(defCurrObj).getRate()%>','<%=currClass.findCurrencyByCurrency(Session("currency")).getRate()%>');"> <%=lang.getTranslated(objListaPayment(k).getKeywordMultilingua())%>&nbsp;<%=strLogo%></li>
						  <%next%>
						</ul>
              <%end if						  
						  
						   Set objModulo = nothing
						   Set objListaPayment = nothing
						   Set objPayment = nothing
              
              If Err.Number<>0 then'
              end if
						  %>
            </div>              
						</div>
						<%
						'************ converto il prezzo in base alla valuta selezionata
						if(hasCurrency) then
							totale_spese_all = currClass.convertCurrency(totSpese+totale_carrello, defCurrObj, Session("currency"))
						else
							totale_spese_all = totSpese+totale_carrello
						end if
						%>   

            <%' *****************************************************		
              ' INIZIO: CODICE GESTIONE ACQUISTO SENZA REGISTRAZIONE
			  ' quando l'utente non Ë gi‡ loggato, quindi non registrato, presento l'opzione per acquistare senza 
              ' registrazione; impostando il campo per l'email e i campi creati per la registrazione utente
              ' e utilizzando id session per username e password

			if(isEmpty(Session("objUtenteLogged")))then%>
				<div class="spese-div">
					<strong><%=lang.getTranslated("frontend.carrello.table.label.buy_noreg")%></strong><br/>
					<select name="buy_noreg" id="buy_noreg">
					  <OPTION VALUE="0" <%if(request("buy_noreg")=0)then response.write("selected") end if%>><%=lang.getTranslated("portal.commons.no")%></OPTION>
					  <OPTION VALUE="1" <%if(request("buy_noreg")=1)then response.write("selected") end if%>><%=lang.getTranslated("portal.commons.yes")%></OPTION>
					</select>
					
					<div id="show_buy_noreg">
					<ul>
					<li><span><%=lang.getTranslated("frontend.area_user.manage.label.email")%> (*)</span></li>
					<li><input type="text" name="noreg_email" id="noreg_email" value="<%=noreg_email%>"/></li><br/>
					</ul>
					
				   <!--******** GESTIONE FIELDS UTENTE PERSONALIZZATI ********-->
			
					<%
					'********** RECUPERO LA LISTA DI FIELD UTENTE DISPONIBILI
					Dim strPrecFieldgroup, strFieldgroup               
					strPrecFieldgroup = ""
							
					On Error Resume next
					Dim userFieldcount, fieldCssClass
					userFieldcount =1
					if(hasUserFields AND Application("show_user_field_on_direct_buy") = 1) then
					  for each k in objListUserField
						Set objField = objListUserField(k)
						fieldCssClass=""
					  
						  if(CInt(objField.getTypeField())=5) then
							fieldCssClass="formFieldMultiple"
						  end if
						  
						  labelForm = objField.getDescription()
						  if not(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())="") then labelForm = lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())
					  
						if(userFieldcount=1) then
						  strFieldgroup = objField.getObjGroup().getDescription()
						  strPrecFieldgroup = strFieldgroup%>
						  <h2><%if not(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)="") then response.write(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)) else response.write(strFieldgroup) end if%></h2>
						  <ul>
						  <li><span><%=labelForm%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></li>
						  <li><%=objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang)%></li><br>
					  <%else
							strFieldgroup = objField.getObjGroup().getDescription()
							if(strFieldgroup = strPrecFieldgroup) then%>
							  <li><span><%=labelForm%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></li>
							  <li><%=objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang)%></li><br>                
							<%else%>
							  </ul>
							  <h2><%if not(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)="") then response.write(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)) else response.write(strFieldgroup) end if%></h2>
							  <ul>
							  <li><span><%=labelForm%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></li>
							  <li><%=objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang)%></li><br>                
							<%strPrecFieldgroup = strFieldgroup
							  end if              
						  end if
							  
						  if(userFieldcount = objListUserField.Count) then response.write("</ul>") end if
							
						userFieldcount=userFieldcount+1
					  next
					end if
			
					if(Err.number<>0) then
					'response.write(Err.description)
					end if        
				  %>
				  <!--******** FINE GESTIONE FIELDS UTENTE PERSONALIZZATI ********-->  					
					</div>				
				</div>
	
				<script>
					<%if(request("buy_noreg") <> "" AND request("buy_noreg")=1)then%>
					$("#show_buy_noreg").show();
					<%else%>
					$("#show_buy_noreg").hide();
					<%end if%>

					$('#buy_noreg').change(function() {
						var val_buy_noreg = $('#buy_noreg').val();
		
						if(val_buy_noreg==1){
							$("#show_buy_noreg").show();

							<%if(Application("show_ship_box") = 1) OR (Application("enable_international_tax_option") = 1) then
								if (applyBills) OR (Application("enable_international_tax_option") = 1) then%>
									$("#card_shipping_address").show();
								<%end if%>		
							<%end if%>
							<%if(Application("show_bills_box") = 1) then%>
								$("#card_bills_address").show();			
							<%end if%>                
						}else{
							$("#show_buy_noreg").hide();

							if ($("#card_shipping_address").length > 0){
								$("#card_shipping_address").hide();
							}
							if ($("#card_shipping_address").length > 0){
								$("#card_bills_address").hide();	
							}
						}
					});
				</script>
			<%else%>
			<input type="hidden" value="0" name="buy_noreg" id="buy_noreg">
			<%end if%>

            <%
			  Dim objCountry 
			  Set objCountry = New CountryClass

              ' *****************************************************		
              ' INIZIO: CODICE GESTIONE SHIPPING ADDRESS
              ' SE L'UTENTE NON E' IN SESSIONE NON PRESENTO NESSUN DATO%>

              <div class="spese-div" id="card_shipping_address">
              <div align="left" onClick="javascript:showHideDiv('divShipCost')"><strong>
			  <%if(Application("enable_international_tax_option") = 1 AND not(applyBills))then%>
			  <%=lang.getTranslated("frontend.carrello.table.label.shipping_address_international_tax")%>
			  <%else%>
			  <%=lang.getTranslated("frontend.carrello.table.label.shipping_address")%>
			  <%end if%>
			  </strong>:<img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="0" hspace="4" width="12" height="16" border="0" align="absmiddle" title="<%=lang.getTranslated("frontend.carrello.table.label.change_ship_address")%>"><br/>
              <%if(hasShipAddress)then
              if(Cint(userIsCompanyClient)=0) then
                userLabelIsCompanyClient = lang.getTranslated("frontend.utenti.detail.table.label.is_private")
              else
                userLabelIsCompanyClient = lang.getTranslated("frontend.utenti.detail.table.label.is_company")
              end if              
              response.write(userName & " " & userSurname & " ("&userLabelIsCompanyClient&") - " & userCfiscVat & " - " &userAddress &" - "&userCity&" ("&userZipCode&") - "&lang.getTranslated("portal.commons.select.option.country."&userCountry)&userStateRegionLabel)
              end if%>
              </div>
              <div id="divShipCost"  <%if not(hasShipAddress)then response.Write("style=""visibility:visible;display:block;""") else response.Write("style=""visibility:hidden;display:none;""") end if%> align="left">
                <br/><div style="float:left;"><strong><%=lang.getTranslated("frontend.carrello.table.label.name")%></strong><br>
                <input type="text" id="ship_name" name="ship_name" value="<%=userName%>"></div>
                <div><strong><%=lang.getTranslated("frontend.carrello.table.label.surname")%></strong><br>
                <input type="text" id="ship_surname" name="ship_surname" value="<%=userSurname%>"></div>
                <div style="float:left;"><strong><%=lang.getTranslated("frontend.carrello.table.label.address")%></strong><br>
                <input type="text" id="ship_address" name="ship_address" value="<%=userAddress%>"></div>
                <div><strong><%=lang.getTranslated("frontend.carrello.table.label.zip_code")%></strong><br>
                <input type="text" id="ship_zip_code" name="ship_zip_code" value="<%=userZipCode%>"></div>
                <div>
                <strong><%=lang.getTranslated("frontend.carrello.table.label.city")%></strong><br>
                <input type="text" id="ship_city" name="ship_city" value="<%=userCity%>"></div>

                <div style="float:left;padding-right:3px;">
                <strong><%=lang.getTranslated("frontend.carrello.table.label.country")%></strong><br>
                <select id="ship_country" name="ship_country">
                <option value=""></option>
                <%On Error Resume next
                Set specialFieldValue = objCountry.findCountryListOnly("2,3")	
                if (Instr(1, typename(specialFieldValue), "Dictionary", 1) > 0) then		    
                for each x in specialFieldValue
                  key =  specialFieldValue(x).getCountryCode()
                  selected = ""
                  if (strComp(key, userCountry, 1) = 0) then selected=" selected" end if%>
                  <option value="<%=key%>" <%=selected%>><%=Server.HTMLEncode(lang.getTranslated("portal.commons.select.option.country."&key))%></option>     
                <%next
                end if
                Set specialFieldValue = nothing
                if(Err.number<>0) then
                end if%>
                </select> 
                </div>
              <div style="padding-bottom:10px;"><strong><%=lang.getTranslated("frontend.carrello.table.label.state_region")%></strong><br>	 
                <select name="ship_state_region" id="ship_state_region">
                <option value=""></option>
                <%
                if(userCountry<>"")then
                  On Error Resume next
                  Set specialFieldValueSR = objCountry.findStateRegionListByCountry(userCountry,"2,3")			
                  if (Instr(1, typename(specialFieldValueSR), "Dictionary", 1) > 0) then		    
                  for each x in specialFieldValueSR
                    key =  specialFieldValueSR(x).getStateRegionCode()
                    selected = ""
                    if (strComp(key, userStateRegion, 1) = 0) then selected=" selected" end if%>
                    <option value="<%=key%>" <%=selected%>><%=Server.HTMLEncode(lang.getTranslated("portal.commons.select.option.country."&key))%></option>     
                  <%next
                  end if
                  Set specialFieldValueSR = nothing
                  if(Err.number<>0) then
                  end if
                end if%>
                </select>	
                </div>

                <div style="float:left;"><strong><%=lang.getTranslated("frontend.carrello.table.label.cfiscvat")%></strong><br>
                <input type="text" id="ship_cfiscvat" name="ship_cfiscvat" value="<%=userCfiscVat%>"></div>
                <div><strong><%=lang.getTranslated("frontend.utenti.detail.table.label.is_company_client")%></strong><br>
                <select name="ship_is_company_client" id="ship_is_company_client">
                <option value="0" <%if (Cint(userIsCompanyClient) = 0) then response.Write("selected")%>><%=lang.getTranslated("frontend.utenti.detail.table.label.is_private")%></option>
                <option value="1" <%if (Cint(userIsCompanyClient) = 1) then response.Write("selected")%>><%=lang.getTranslated("frontend.utenti.detail.table.label.is_company")%></option>
                </select></div>
				<script>
				// se Ë stata modificato country al refresh ripristino i valori dei campi utente personalizzati
				<%
				if(hasUserFields AND Application("show_user_field_on_direct_buy") = 1) then%>
					$(document).ready(function() {
					<%for each k in objListUserField
						Set objField = objListUserField(k)

						select Case objField.getTypeField()
						Case 6,7%>
							var fieldTmp = document.form_insert_carrello.<%=objUserField.getFieldPrefix()&objField.getID()%>;
							var fieldTmpVal = "<%=request(objUserField.getFieldPrefix()&objField.getID())%>";

							if (fieldTmp){
								if(fieldTmp != null){
									var splittedVals = fieldTmpVal.split(",");
									if(fieldTmp.length == null){
										for (var j=0; j < splittedVals.length; j++){
											if (fieldTmp.value==splittedVals[j]){
												fieldTmp.checked=true;
												break;
											}
										}
									}else{
										for (var i=0; i < fieldTmp.length; i++){
											for (var j=0; j < splittedVals.length; j++){
												if (fieldTmp[i].value==splittedVals[j]){
													fieldTmp[i].checked=true;
												}
											}
										}
									}
								}
							}
						<%Case Else%>
							$('#<%=objUserField.getFieldPrefix()&objField.getID()%>').val('<%=request(objUserField.getFieldPrefix()&objField.getID())%>');			
						<%End Select%>
			
						<%Set objField = nothing
					next%>
					});
				<%end if

				if(Application("enable_international_tax_option") = 1) then%>	

				$('#ship_country').change(function() {
					$('#prodotto-totale').hide();					

					var query_string = "?ship_country="+escape($('#ship_country').val())
					query_string+="&ship_state_region="
					query_string+="&buy_noreg="+escape($('#buy_noreg').val());

					query_string+="&noreg_email="+escape($('#noreg_email').val());

					query_string+="&ship_name="+escape($('#ship_name').val());
					query_string+="&ship_surname="+escape($('#ship_surname').val());
					query_string+="&ship_cfiscvat="+escape($('#ship_cfiscvat').val());
					query_string+="&ship_address="+escape($('#ship_address').val());
					query_string+="&ship_zip_code="+escape($('#ship_zip_code').val());
					query_string+="&ship_city="+escape($('#ship_city').val());
					query_string+="&ship_is_company_client="+escape($('#ship_is_company_client').val());

					<%if(Application("show_bills_box") = 1)then%>
					query_string+="&bills_name="+escape($('#bills_name').val());
					query_string+="&bills_surname="+escape($('#bills_surname').val());
					query_string+="&bills_cfiscvat="+escape($('#bills_cfiscvat').val());
					query_string+="&bills_address="+escape($('#bills_address').val());
					query_string+="&bills_zip_code="+escape($('#bills_zip_code').val());
					query_string+="&bills_city="+escape($('#bills_city').val());
					query_string+="&bills_country="+escape($('#bills_country').val());
					query_string+="&bills_state_region="+escape($('#bills_state_region').val());
					<%end if

					if(hasUserFields AND Application("show_user_field_on_direct_buy") = 1) then
						for each k in objListUserField
							Set objField = objListUserField(k)

							select Case objField.getTypeField()
							Case 6,7%>
								var fieldTmpS = document.form_insert_carrello.<%=objUserField.getFieldPrefix()&objField.getID()%>;
								var fieldTmpValS = "";
								if (fieldTmpS){
									if(fieldTmpS != null){
										if(fieldTmpS.length == null){
											if (fieldTmpS.checked){
												fieldTmpValS = fieldTmpValS + fieldTmpS.value + ",";
											}
										}else{
											for (var i=0; i < fieldTmpS.length; i++){
												if (fieldTmpS[i].checked){
													fieldTmpValS = fieldTmpValS + fieldTmpS[i].value + ",";
												}
											}
										}
										fieldTmpValS = fieldTmpValS.substring(0, fieldTmpValS.lastIndexOf(','));
									}
								}
								query_string+="&<%=objUserField.getFieldPrefix()&objField.getID()%>="+escape(fieldTmpValS);	
							<%Case Else%>
								query_string+="&<%=objUserField.getFieldPrefix()&objField.getID()%>="+escape($('#<%=objUserField.getFieldPrefix()&objField.getID()%>').val());				
							<%End Select%>					
							<%Set objField = nothing
						next
					end if
					%>

					var url = '<%=Application("baseroot") &Application("dir_upload_templ")&"shopping-card/carrello2.asp"%>'+query_string;

					window.location.href = url;
				});

				$('#ship_state_region').change(function() {
					$('#prodotto-totale').hide();

					var query_string = "?ship_country="+escape($('#ship_country').val())
					query_string+="&ship_state_region="+escape($('#ship_state_region').val())
					query_string+="&buy_noreg="+escape($('#buy_noreg').val());

					query_string+="&noreg_email="+escape($('#noreg_email').val());

					query_string+="&ship_name="+escape($('#ship_name').val());
					query_string+="&ship_surname="+escape($('#ship_surname').val());
					query_string+="&ship_cfiscvat="+escape($('#ship_cfiscvat').val());
					query_string+="&ship_address="+escape($('#ship_address').val());
					query_string+="&ship_zip_code="+escape($('#ship_zip_code').val());
					query_string+="&ship_city="+escape($('#ship_city').val());
					query_string+="&ship_is_company_client="+escape($('#ship_is_company_client').val());

					<%if(Application("show_bills_box") = 1)then%>
					query_string+="&bills_name="+escape($('#bills_name').val());
					query_string+="&bills_surname="+escape($('#bills_surname').val());
					query_string+="&bills_cfiscvat="+escape($('#bills_cfiscvat').val());
					query_string+="&bills_address="+escape($('#bills_address').val());
					query_string+="&bills_zip_code="+escape($('#bills_zip_code').val());
					query_string+="&bills_city="+escape($('#bills_city').val());
					query_string+="&bills_country="+escape($('#bills_country').val());
					query_string+="&bills_state_region="+escape($('#bills_state_region').val());
					<%end if
					
					if(hasUserFields AND Application("show_user_field_on_direct_buy") = 1) then
						for each k in objListUserField
							Set objField = objListUserField(k)							

							select Case objField.getTypeField()
							Case 6,7%>
								var fieldTmpB = document.form_insert_carrello.<%=objUserField.getFieldPrefix()&objField.getID()%>;
								var fieldTmpValB = "";
								if (fieldTmpB){
									if(fieldTmpB != null){
										if(fieldTmpB.length == null){
											if (fieldTmpB.checked){
												fieldTmpValB = fieldTmpValB + fieldTmpB.value + ",";
											}
										}else{
											for (var i=0; i < fieldTmpB.length; i++){
												if (fieldTmpB[i].checked){
													fieldTmpValB = fieldTmpValB + fieldTmpB[i].value + ",";
												}
											}
										}
										fieldTmpValB = fieldTmpValB.substring(0, fieldTmpValB.lastIndexOf(','));
									}
								}
								query_string+="&<%=objUserField.getFieldPrefix()&objField.getID()%>="+escape(fieldTmpValB);	
							<%Case Else%>
								query_string+="&<%=objUserField.getFieldPrefix()&objField.getID()%>="+escape($('#<%=objUserField.getFieldPrefix()&objField.getID()%>').val());				
							<%End Select%>					
							<%Set objField = nothing
						next
					end if
					%>

					var url = '<%=Application("baseroot") &Application("dir_upload_templ")&"shopping-card/carrello2.asp"%>'+query_string;

					window.location.href = url;
				});

				$('#ship_is_company_client').change(function() {
					$('#prodotto-totale').hide();

					var query_string = "?ship_country="+escape($('#ship_country').val())
					query_string+="&ship_state_region="+escape($('#ship_state_region').val())
					query_string+="&buy_noreg="+escape($('#buy_noreg').val());

					query_string+="&noreg_email="+escape($('#noreg_email').val());

					query_string+="&ship_name="+escape($('#ship_name').val());
					query_string+="&ship_surname="+escape($('#ship_surname').val());
					query_string+="&ship_cfiscvat="+escape($('#ship_cfiscvat').val());
					query_string+="&ship_address="+escape($('#ship_address').val());
					query_string+="&ship_zip_code="+escape($('#ship_zip_code').val());
					query_string+="&ship_city="+escape($('#ship_city').val());
					query_string+="&ship_is_company_client="+escape($('#ship_is_company_client').val());

					<%if(Application("show_bills_box") = 1)then%>
					query_string+="&bills_name="+escape($('#bills_name').val());
					query_string+="&bills_surname="+escape($('#bills_surname').val());
					query_string+="&bills_cfiscvat="+escape($('#bills_cfiscvat').val());
					query_string+="&bills_address="+escape($('#bills_address').val());
					query_string+="&bills_zip_code="+escape($('#bills_zip_code').val());
					query_string+="&bills_city="+escape($('#bills_city').val());
					query_string+="&bills_country="+escape($('#bills_country').val());
					query_string+="&bills_state_region="+escape($('#bills_state_region').val());
					<%end if
					
					if(hasUserFields AND Application("show_user_field_on_direct_buy") = 1) then
						for each k in objListUserField
							Set objField = objListUserField(k)							

							select Case objField.getTypeField()
							Case 6,7%>
								var fieldTmpB = document.form_insert_carrello.<%=objUserField.getFieldPrefix()&objField.getID()%>;
								var fieldTmpValB = "";
								if (fieldTmpB){
									if(fieldTmpB != null){
										if(fieldTmpB.length == null){
											if (fieldTmpB.checked){
												fieldTmpValB = fieldTmpValB + fieldTmpB.value + ",";
											}
										}else{
											for (var i=0; i < fieldTmpB.length; i++){
												if (fieldTmpB[i].checked){
													fieldTmpValB = fieldTmpValB + fieldTmpB[i].value + ",";
												}
											}
										}
										fieldTmpValB = fieldTmpValB.substring(0, fieldTmpValB.lastIndexOf(','));
									}
								}
								query_string+="&<%=objUserField.getFieldPrefix()&objField.getID()%>="+escape(fieldTmpValB);	
							<%Case Else%>
								query_string+="&<%=objUserField.getFieldPrefix()&objField.getID()%>="+escape($('#<%=objUserField.getFieldPrefix()&objField.getID()%>').val());				
							<%End Select%>					
							<%Set objField = nothing
						next
					end if
					%>

					var url = '<%=Application("baseroot") &Application("dir_upload_templ")&"shopping-card/carrello2.asp"%>'+query_string;

					window.location.href = url;
				});
				
				<%else%>

				$('#ship_country').change(function() {
					var type_val_ch = $('#ship_country').val();
					var query_string = "field_val="+escape(type_val_ch);

					$.ajax({
						async: true,
						type: "GET",
						cache: false,
						url: "<%=Application("baseroot") & "/common/include/ajaxstateregionlistupdate.asp"%>",
						data: query_string,
						success: function(response) {
							//alert("response: "+response);
							$("select#ship_state_region").empty();
							$("select#ship_state_region").append($("<option></option>").attr("value","").text(""));
							$("select#ship_state_region").append(response);
						},
						error: function() {
							$("select#ship_state_region").empty();
							$("select#ship_state_region").append($("<option></option>").attr("value","").text(""));
						}
					});		
				});
				<%end if%>
				</script>
              </div>            
            </div>
	
			<script>
				$("#card_shipping_address").hide();

				<%if(Application("show_ship_box") = 1) OR (Application("enable_international_tax_option") = 1) OR (request("buy_noreg") = 1 AND ((Application("show_ship_box") = 1) OR (Application("enable_international_tax_option") = 1))) then
					if ((not(isEmpty(Session("objUtenteLogged"))) OR (request("buy_noreg") = 1))  AND (applyBills)) OR (Application("enable_international_tax_option") = 1 AND request("buy_noreg") = 1) OR (Application("enable_international_tax_option") = 1 AND not(isEmpty(Session("objUtenteLogged")))) then%>
						$("#card_shipping_address").show();
					<%end if
				end if%>
			</script>


            <%			
			  Set objListUserField = nothing
			  Set objUserField = nothing

			  ' *****************************************************		
              ' INIZIO: CODICE GESTIONE BILLS ADDRESS
              ' SE L'UTENTE NON E' IN SESSIONE NON PRESENTO NESSUN DATO
              '*** Ë necessario che sia sempre visibile, anche per gli ordini con prodotti solo scaricabili;
              '*** senza i dati di fatturazione non si riesce a generare una fattura fiscalmente valida%>

              <div class="spese-div" id="card_bills_address">
              <div align="left" onClick="javascript:showHideDiv('divBillsCost')"><strong><%=lang.getTranslated("frontend.carrello.table.label.bills_address")%></strong>:<img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="0" hspace="4" width="12" height="16" border="0" align="absmiddle" title="<%=lang.getTranslated("frontend.carrello.table.label.change_bills_address")%>"><br/>
              <%if(hasBillsAddress)then
              response.write(buserName & " " & buserSurname & " - " & buserCfiscVat & " - " &buserAddress &" - "&buserCity&" ("&buserZipCode&") - "&lang.getTranslated("portal.commons.select.option.country."&buserCountry)&buserStateRegionLabel)
              end if%>
              </div>
              <div id="divBillsCost"  <%if not(hasBillsAddress)then response.Write("style=""visibility:visible;display:block;""") else response.Write("style=""visibility:hidden;display:none;""") end if%> align="left">
                <br/><div style="float:left;"><strong><%=lang.getTranslated("frontend.carrello.table.label.name")%></strong><br>
                <input type="text" id="bills_name" name="bills_name" value="<%=buserName%>"></div>
                <div><strong><%=lang.getTranslated("frontend.carrello.table.label.surname")%></strong><br>
                <input type="text" id="bills_surname" name="bills_surname" value="<%=buserSurname%>"></div>
                <div style="float:left;"><strong><%=lang.getTranslated("frontend.carrello.table.label.address")%></strong><br>
                <input type="text" id="bills_address" name="bills_address" value="<%=buserAddress%>"></div>
                <div><strong><%=lang.getTranslated("frontend.carrello.table.label.zip_code")%></strong><br>
                <input type="text" id="bills_zip_code" name="bills_zip_code" value="<%=buserZipCode%>"></div>
                <div>
                <strong><%=lang.getTranslated("frontend.carrello.table.label.city")%></strong><br>
                <input type="text" id="bills_city" name="bills_city" value="<%=buserCity%>"></div>

                <div style="float:left;padding-right:3px;">
                <strong><%=lang.getTranslated("frontend.carrello.table.label.country")%></strong><br>
                <select id="bills_country" name="bills_country">
                <option value=""></option>
                <%On Error Resume next
                Set specialFieldValueB = objCountry.findCountryListOnly("2,3")
                if (Instr(1, typename(specialFieldValueB), "Dictionary", 1) > 0) then			    
                for each x in specialFieldValueB
                  key =  specialFieldValueB(x).getCountryCode()
                  selected = ""
                  if (strComp(key, buserCountry, 1) = 0) then selected=" selected" end if%>
                  <option value="<%=key%>" <%=selected%>><%=Server.HTMLEncode(lang.getTranslated("portal.commons.select.option.country."&key))%></option>     
                <%next
                end if
                Set specialFieldValueB = nothing
                if(Err.number<>0) then
                end if%>
                </select>  
                </div>  
                <div style="padding-bottom:10px;">
                <strong><%=lang.getTranslated("frontend.carrello.table.label.state_region")%></strong><br>	 
                <select name="bills_state_region" id="bills_state_region">
                <option value=""></option>
                <%
                if(buserCountry<>"")then
                  On Error Resume next
                  Set specialFieldValueBSR = objCountry.findStateRegionListByCountry(buserCountry,"2,3")	
                  if (Instr(1, typename(specialFieldValueBSR), "Dictionary", 1) > 0) then	    
                  for each x in specialFieldValueBSR
                    key =  specialFieldValueBSR(x).getStateRegionCode()
                    selected = ""
                    if (strComp(key, buserStateRegion, 1) = 0) then selected=" selected" end if%>
                    <option value="<%=key%>" <%=selected%>><%=Server.HTMLEncode(lang.getTranslated("portal.commons.select.option.country."&key))%></option>     
                  <%next
                  end if
                  Set specialFieldValueBSR = nothing					
                  if(Err.number<>0) then
                  end if
                end if%>
                </select>	  
                </div>
                <div><strong><%=lang.getTranslated("frontend.carrello.table.label.cfiscvat")%></strong><br>
                <input type="text" id="bills_cfiscvat" name="bills_cfiscvat" value="<%=buserCfiscVat%>"></div>
				<script>
				$('#bills_country').change(function() {
					var type_val_ch = $('#bills_country').val();
					var query_string = "field_val="+escape(type_val_ch);

					$.ajax({
						async: true,
						type: "GET",
						cache: false,
						url: "<%=Application("baseroot") & "/common/include/ajaxstateregionlistupdate.asp"%>",
						data: query_string,
						success: function(response) {
							//alert("response: "+response);
							$("select#bills_state_region").empty();
							$("select#bills_state_region").append($("<option></option>").attr("value","").text(""));
							$("select#bills_state_region").append(response);
						},
						error: function() {
							$("select#bills_state_region").empty();
							$("select#bills_state_region").append($("<option></option>").attr("value","").text(""));
						}
					});		
				});
				</script> 
              </div>
              <%Set objBills = nothing
				Set objCountry = nothing%>             
            </div>
	
			<script>
				$("#card_bills_address").hide();

				<%if(Application("show_bills_box") = 1) then
					if not(isEmpty(Session("objUtenteLogged"))) OR (request("buy_noreg") = 1) then%>
						$("#card_bills_address").show();
					<%end if
				end if%>
			</script>
                    
            
            <div id="spese-totale">
			  <%=lang.getTranslated("frontend.carrello.table.label.payment_commission")%>: <strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<span class="payment_commission"><%=FormatNumber(payment_commission, 2,-1)%></span></strong><br/><br/> 
              <%=lang.getTranslated("frontend.carrello.table.label.totale_ordine")%>: <strong><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<span class="ord_total"><%=FormatNumber(totale_spese_all, 2,-1)%></span></strong>
			<br/><span id="currency"><%=lang.getTranslated("frontend.carrello.table.label.totale_ordine.currency_transaction1")%>&nbsp;<%=lang.getTranslated("backend.currency.keyword.label."&defCurrObj)%>&nbsp;<%=lang.getTranslated("frontend.carrello.table.label.totale_ordine.currency_transaction2")%>:&nbsp;<%if(lang.getTranslated("backend.currency.symbol.label."&defCurrObj) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&defCurrObj)) else response.write(defCurrObj) end if%>&nbsp;<span class="ord_total_def_curr"><%=FormatNumber(totSpese+totale_carrello, 2,-1)%></span></span>
			</div>

            <div id="prodotto-totale">
						<a href="javascript:sendCarrello(<%if(applyBills) then response.write("1") else response.write("0") end if%>);"><%=lang.getTranslated("frontend.carrello.table.label.confirm_order")%></a>&nbsp;&nbsp;
						<a href="<%=Application("baseroot") &Application("dir_upload_templ")&"shopping-card/DeleteCarrello.asp?id_carrello_to_delete=" & id_carrello%>"><%=lang.getTranslated("frontend.carrello.table.label.cancel_card")%></a>
						</div>
            </form>			
					</div>
				</div>
				<%Set objListaCarrello = nothing			
			else
				response.write("<br><br><p align='center'>"&lang.getTranslated("frontend.carrello.table.label.empty_card")&"</p>")
			end if
		else			
			response.write("<br><br><p align='center'>"&lang.getTranslated("frontend.carrello.table.label.empty_card")&"</p>")
		end if%>