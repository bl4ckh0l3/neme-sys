		<div id="torna"><a href="<%=Application("baseroot")&"/common/include/Controller.asp?gerarchia="&strGerarchia%>"><%=lang.getTranslated("frontend.template_prodotto.table.label.back_to_list")%></a>
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
		</div>
		<%    
		response.Write("<h1><br/>"&objCurrentProdotto.findFieldTranslation(1 ,lang.getLangCode(),1) & "</h1>")
		response.Write("<p>"&objCurrentProdotto.findFieldTranslation(3 ,lang.getLangCode(),1) & "</p>")%>		
    
    <form action="<%=Application("baseroot")&Application("dir_upload_templ")&"shopping-card/ManageCarrello.asp"%>" method="post" name="form_add_to_carrello" enctype="multipart/form-data">
    
    <%if(Application("disable_ecommerce") = 0) OR (Application("disable_ecommerce") = 1) then%>
    <%		
		Dim numPrezzoReal, numPrezzoOld
		numPrezzoReal = objCurrentProdotto.getPrezzo() 
    numPrezzoOld = numPrezzoReal 
    discountPercent = 0
            
    if(hasGroup) then
      On Error Resume Next
      discountPercent = objSelMargin.getDiscountPercentual(CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objCurrentProdotto.getsconto()),CDbl(scontoCliente))
      numPrezzoReal = objSelMargin.getAmount(numPrezzoReal,CDbl(objSelMargin.getMargin()),CDbl(objSelMargin.getDiscount()),objSelMargin.isApplyProdDiscount(),objSelMargin.isApplyUserDiscount(),CDbl(objCurrentProdotto.getsconto()),CDbl(scontoCliente))
      if(Err.number <>0) then
      end if
    else
      if(objCurrentProdotto.hasSconto() AND (not(hasSconto) OR (hasSconto AND Application("manage_sconti") = 1))) then
        numPrezzoReal = objCurrentProdotto.getPrezzoScontato()
        discountPercent = CDbl(objCurrentProdotto.getsconto())
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
    
    Set objTasse = new TaxsClass
    
    
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
        Set groupProdTax = objCurrentProdotto.getTaxGroupObj(objCurrentProdotto.getTaxGroup()) 
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
      if not(isNull(objCurrentProdotto.getIDTassaApplicata())) AND not(objCurrentProdotto.getIDTassaApplicata() = "") then
        numPrezzoReal = numPrezzoReal+objCurrentProdotto.getImportoTassa(numPrezzoReal)
        numPrezzoOld = numPrezzoOld+objCurrentProdotto.getImportoTassa(numPrezzoOld)
        descTassa = objTasse.findTassaByID(objCurrentProdotto.getIDTassaApplicata()).getDescrizioneTassa()
      end if
    end if
    
    if(descTassa<>"") then descTassa = "&nbsp;&nbsp;("&descTassa&")" end if       
    
    
		Set objTasse = nothing   
    
    
		'************ converto il prezzo in base alla valuta selezionata
		if not(defCurrObj="") AND not(thisCurrObj="") then
			numPrezzoOld = currClass.convertCurrency(numPrezzoOld, defCurrObj, thisCurrObj)
			numPrezzoReal = currClass.convertCurrency(numPrezzoReal, defCurrObj, thisCurrObj)
		end if
    %>
    <h3>
		<%if(discountPercent > 0) then%><span class="testoBarrato"><%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(numPrezzoOld, 2,-1)&descTassa%></span><br/><%end if%>
		<%if(lang.getTranslated("backend.currency.symbol.label."&Session("currency")) <> "") then response.write(lang.getTranslated("backend.currency.symbol.label."&Session("currency"))) else response.write(Session("currency")) end if%>&nbsp;<%=FormatNumber(numPrezzoReal, 2,-1) &descTassa & "</h3>"%>
		<%if(discountPercent > 0) then response.Write("<h3><strong>"&lang.getTranslated("frontend.template_prodotto.table.label.sconto") & " " & discountPercent & "%</strong></h3>") end if%>		
		<%if(objCurrentProdotto.getAttivo() = 0) then response.Write("<h3>"&lang.getTranslated("frontend.template_prodotto.table.alt.non_disponibile") & "</h3>") end if%>
		<input type="hidden" value="<%=objCurrentProdotto.getIDProdotto()%>" name="id_prodotto">
		<input type="hidden" value="add" name="operation">
    
		<%
		On Error Resume Next
		hasProdFields=false
		
		Set objListProdField = objProdField.getListProductField4ProdActive(objCurrentProdotto.getIDProdotto())
		
		if(objListProdField.Count > 0)then
			hasProdFields = true
		end if
		
		if(Err.number <> 0) then
			hasProdFields = false
		end if	
		
		if(hasProdFields)then	
			
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
						response.write(labelForm & ":&nbsp;" &objProdField.renderProductFieldHTML(objField,fieldCssClass, "", objCurrentProdotto.getIDProdotto(), "",lang,1,objField.getEditable()) & "<br/>")%>
					<%
					else
						valueTmp = objField.getSelValue()
						response.write(labelForm & ":&nbsp;" & valueTmp & "<br/>")
					end if			
				Case 3,4,5,6						
					if(CInt(objField.getTypeField())=4) then
						fieldCssClass="formFieldMultiple"
					end if
					response.write(labelForm & ":&nbsp;" &objProdField.renderProductFieldHTML(objField,fieldCssClass, "", objCurrentProdotto.getIDProdotto(), "",lang,1,objField.getEditable()) & "<br/>")
        Case 7
          fieldValueMatch = objProdField.findFieldMatchValue(k,objCurrentProdotto.getIDProdotto())
          response.write(objProdField.renderProductFieldHTML(objField,fieldCssClass, "", objCurrentProdotto.getIDProdotto(), fieldValueMatch,lang,1,objField.getEditable()))
				Case 8
					fieldCssClass="formFieldTXTMedium"
					if(objField.getEditable()="1")then
						response.write(labelForm & ":&nbsp;" &objProdField.renderProductFieldHTML(objField,fieldCssClass, "", objCurrentProdotto.getIDProdotto(), "",lang,1,objField.getEditable()) & "<br/>")%>
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
            var cloptions<%=objProdField.getFieldPrefix()&objField.getID()%> = {
            width:280,	// width not including margins, borders or padding
            height:200,	// height not including margins, borders or padding
            controls:"bold italic underline strikethrough subscript superscript | font size style | color highlight removeformat | bullets numbering | alignleft center alignright justify | rule | cut copy paste | image",	// controls to add to the toolbar
            }
						</script>
						<%response.write(labelForm & ":&nbsp;" &objProdField.renderProductFieldHTML(objField,fieldCssClass, "", objCurrentProdotto.getIDProdotto(), "",lang,1,objField.getEditable()) & "<br/>")%>
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
	end if
%>
    
		<%if(not(isNull(objCommento.findCommentiByIDElement(objCurrentProdotto.getIDProdotto(),2,1)))) AND (Instr(1, typename(objCommento.findCommentiByIDElement(objCurrentProdotto.getIDProdotto(),2,1)), "dictionary", 1) > 0) then%>
		<a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popupComments.asp?id_element="&objCurrentProdotto.getIDProdotto()&"&element_type=2&active=1"%>','popupallegati',400,400,100,100);" title="<%=lang.getTranslated("frontend.template_prodotto.table.alt.see_comments")%>"><img src="<%=Application("baseroot")&"/common/img/ico-commenti.png"%>" hspace="0" vspace="0" border="0"></a>
		<%end if%>
		
    <%if(Application("disable_ecommerce") = 0) then%>
      <%if(objCurrentProdotto.getAttivo() = 1) then%>
      <a href="javascript:addToCarrello(document.form_add_to_carrello);" title="<%=lang.getTranslated("frontend.template_prodotto.table.alt.put_on_carrello")%>"><img src="<%=Application("baseroot")&"/common/img/cart_add.png"%>" hspace="0" vspace="0" border="0"></a>
      <%else%>
      <img src="<%=Application("baseroot")&"/common/img/ico-carrello.png"%>" hspace="0" vspace="0" border="0" alt="<%=lang.getTranslated("frontend.template_prodotto.table.alt.non_disponibile")%>">
      <%end if%>
      
      <%'GESTISCO LA QUANTITA' SELEZIONABILE								
      if(objCurrentProdotto.getQtaDisp() = Application("unlimited_key")) then%>
        <input type="text" name="qta_prodotto" value="" onkeypress="javascript:return isInteger(event);" class="formFieldTXTShort">
      <%else%>
        <!--<select name="qta_prodotto" class="formFieldSelect">
        <option value=""></option>
        <%'for statiOrderCount = 1 to objCurrentProdotto.getQtaDisp()%>
          <option value="<%'=statiOrderCount%>" ><%'=statiOrderCount%></option>
        <%'next%>
        </select>-->
        <input type="text" name="qta_prodotto" id="qta_prodotto" value="" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);" onblur="javascript:checkMaxQtaProd(<%=objCurrentProdotto.getQtaDisp()%>,this);">&nbsp;<%=lang.getTranslated("frontend.template_prodotto.table.label.product_disp")&"&nbsp;"&objCurrentProdotto.getQtaDisp()%>&nbsp;&nbsp;	
      <%end if%>
    <%end if%>
		</form>
<br/><br/>        
		<%
		On Error Resume next
		Set objRelationsProd = objCurrentProdotto.getRelationPerProdotto(id_prodotto)
		if (strComp(typename(objRelationsProd), "Dictionary", 1) = 0) then%>
			<div style="display:block " align="left">
			<h2><%=lang.getTranslated("frontend.template_prodotto.table.label.related_prod")%></h2>
			<%
			Set objCat = new CategoryClass
			Set objPageTempl = new Page4TemplateClass

			counter = 1
			for each k in objRelationsProd
				Set objRelProd = objRelationsProd(k)

				On Error Resume Next
					Set objFilesRelProd = objRelProd.getFileXProdotto()	
				If(Err.number <> 0) then
					objFilesRelProd = null
				end if%>

				<%if(counter MOD 5 = 0)then%><div id="clear"></div><%end if%>
				<div id="prodotto-immagine">
				<%
					Dim urlRelProd, gerCatRelProd
					urlRelProd = "#"
					
					gerCatRelProd = objCurrentProdotto.getGerCatProd4Relation(objRelProd.getIDProdotto())
					numPageTempl = objPageTempl.getMaxNumPageByIDTemplate(objCat.findExsitingCategoriaByGerarchia(gerCatRelProd).getIDTemplate())
					
					if(gerCatRelProd<> "" AND numPageTempl <> "") then
						urlRelProd = Application("baseroot") & "/common/include/Controller.asp?gerarchia="&gerCatRelProd&"&id_prodotto="&objRelProd.getIDProdotto()&"&page=&modelPageNum="&numPageTempl
					end if

					if not(isNull(objFilesRelProd)) then%>
						<%Dim hasNotSmallImg
						hasNotSmallImg = true			
						for each xObjFile in objFilesRelProd
							Set objFileXProdotto = objFilesRelProd(xObjFile)
							iTypeFile = objFileXProdotto.getFileTypeLabel()
							if(Cint(iTypeFile) = 1) then%>	
								<a href="<%=urlRelProd%>" title="<%=objRelProd.findFieldTranslation(1 ,lang.getLangCode(),1)%>"><img src="<%=Application("dir_upload_prod")&objFileXProdotto.getFilePath()%>" alt="<%=objRelProd.findFieldTranslation(1 ,lang.getLangCode(),1)%>" width="100" height="100" align="top" /></a>
								<%hasNotSmallImg = false
								Exit for
							end if
							Set objFileXProdotto = nothing	
						next		
						if(hasNotSmallImg) then%>
						<a href="<%=urlRelProd%>" title="<%=objRelProd.findFieldTranslation(1 ,lang.getLangCode(),1)%>"><img width="100" height="100" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" alt="<%=objRelProd.findFieldTranslation(1 ,lang.getLangCode(),1)%>" hspace="0" vspace="0" border="0" align="top"></a>
						<%end if
						Set objFilesRelProd = nothing
					else%>
						<a href="<%=urlRelProd%>" title="<%=objRelProd.findFieldTranslation(1 ,lang.getLangCode(),1)%>"><img width="100" height="100" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" alt="<%=objRelProd.findFieldTranslation(1 ,lang.getLangCode(),1)%>" hspace="0" vspace="0" border="0" align="top"></a>
					<%end if%>
				</div>
				<%Set objRelProd = nothing				
				counter = counter +1
			next
			Set objCat = nothing
			Set objPageTempl = nothing%>
			</div>
		<%end if

		Set objRelationsProd = nothing

		if(Err.number<>0) then
		'response.write(Err.description)
		end if	 
    
Set objCurrentProdotto = nothing%>				