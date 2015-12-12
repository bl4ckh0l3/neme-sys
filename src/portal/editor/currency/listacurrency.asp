<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function deleteCurrency(id_objref,row,refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.currency.lista.js.alert.delete_currency")%>?")){
		ajaxDeleteItem(id_objref,"currency",row,refreshrows);
	}
}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LCY"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
			<div id="ajaxresp" align="center" style="background-color:#FFFF00; border:1px solid #000000; color:#000000; display:none;"></div>
			<table border="0" cellpadding="0" cellspacing="0" class="principal">
              <tr> 
				  <th colspan="2">&nbsp;</th>
				  <th><%=UCase(langEditor.getTranslated("backend.currency.lista.table.header.descrizione"))%></th>
				  <th><%=UCase(langEditor.getTranslated("backend.currency.lista.table.header.valore"))%></th>
				  <th><%=UCase(langEditor.getTranslated("backend.currency.lista.table.header.abilitato"))%></th>
				  <th><%=UCase(langEditor.getTranslated("backend.currency.lista.table.header.default"))%></th>
				  <th><%=UCase(langEditor.getTranslated("backend.currency.lista.table.header.dta_riferimento"))%></th>
				  <th><%=UCase(langEditor.getTranslated("backend.currency.lista.table.header.dta_inserimento"))%></th>
              </tr> 
				<%
				On Error Resume Next
				Dim hasCurrency
				hasCurrency = false
				Set objListaCurrency = objCurrency.getListaCurrency(null,null,null)
				hasCurrency = true				
				
				if Err.number <> 0 then
					hasCurrency = false
				end if
				
				if(hasCurrency) then			
					Dim intCount
					intCount = 0
					
					Dim newsCounter, iIndex, objTmpCurr, objTmpCurrKey, FromCurr, ToCurr, Diff
					iIndex = objListaCurrency.Count
					FromCurr = ((numPage * itemsXpage) - itemsXpage)
					Diff = (iIndex - ((numPage * itemsXpage)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToCurr = iIndex - Diff
					
					totPages = iIndex\itemsXpage
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
						totPages = totPages +1	
					end if		
				
					Dim styleRow, styleRow2
					styleRow2 = "table-list-on"							
							
					objTmpCurr = objListaCurrency.Items
					objTmpCurrKey=objListaCurrency.Keys	
					for newsCounter = FromCurr to ToCurr
						styleRow = "table-list-off"
						if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>
						<form action="<%=Application("baseroot") & "/editor/currency/InserisciCurrency.asp"%>" method="post" name="form_lista_<%=intCount%>">
						<input type="hidden" value="<%=objTmpCurrKey(newsCounter)%>" name="id_currency">
						<input type="hidden" value="" name="delete_currency">
						<input type="hidden" value="LCY" name="cssClass">		
						</form> 
						<tr class="<%=styleRow%>" id="tr_delete_list_<%=intCount%>">
							<%Set objTmpCurr0 = objTmpCurr(newsCounter)%>	
							<td align="center" width="25"><a href="javascript:document.form_lista_<%=intCount%>.submit();"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.currency.lista.table.alt.modify_currency")%>" hspace="2" vspace="0" border="0"></a></td>
							<td align="center" width="25">
								<%if(objTmpCurr0.getDefault() = "0") then%>
								<a href="javascript:deleteCurrency(<%=objTmpCurrKey(newsCounter)%>, 'tr_delete_list_<%=intCount%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.currency.lista.table.alt.delete_currency")%>" hspace="2" vspace="0" border="0"></a>
								<%end if%>
							</td>						
							<td><b><%=objTmpCurr0.getCurrency()%></b>&nbsp;<%if(langEditor.getTranslated("backend.currency.keyword.label."&objTmpCurr0.getCurrency()) <> "") then response.write("("&langEditor.getTranslated("backend.currency.keyword.label."&objTmpCurr0.getCurrency())&")") end if%></td>
							<td><%=FormatNumber(objTmpCurr0.getRate(),4,-1)%></td>
							<td width="8%">
							<div class="ajax" id="view_attivo_<%=intCount%>" onmouseover="javascript:showHide('view_attivo_<%=intCount%>','edit_attivo_<%=intCount%>','attivo_<%=intCount%>',500, true);">
							<%
							Select Case objTmpCurr0.getActive()
							Case 0
								response.write(langEditor.getTranslated("backend.commons.no"))
							Case 1
								response.write(langEditor.getTranslated("backend.commons.yes"))
							Case Else
							End Select%>
							</div>
							<div class="ajax" id="edit_attivo_<%=intCount%>">
							<select name="attivo" class="formfieldAjaxSelect" id="attivo_<%=intCount%>" onblur="javascript:updateField('edit_attivo_<%=intCount%>','view_attivo_<%=intCount%>','attivo_<%=intCount%>','currency',<%=objTmpCurr0.getID()%>,2,<%=intCount%>);">
							<OPTION VALUE="0" <%if (strComp("0", objTmpCurr0.getActive(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
							<OPTION VALUE="1" <%if (strComp("1", objTmpCurr0.getActive(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
							</SELECT>	
							</div>
							<script>
							$("#edit_attivo_<%=intCount%>").hide();
							</script>
							</td>
							<td>
							<%
							Select Case objTmpCurr0.getDefault()
							Case 0
								response.write(langEditor.getTranslated("backend.commons.no"))
							Case 1
								response.write(langEditor.getTranslated("backend.commons.yes"))
							Case Else
							End Select%>
							</td>
							<td><%=FormatDateTime(objTmpCurr0.getDtaRefer(),2)%></td>
							<td><%=FormatDateTime(objTmpCurr0.getDtaInsert(),2)%>&nbsp;<%=DatePart("h",objTmpCurr0.getDtaInsert())%>:<%=DatePart("n",objTmpCurr0.getDtaInsert())%></td>
						</tr>		
						<%						
						Set objTmpCurr0 = nothing
						intCount = intCount +1
					next
					Set objListaCurrency = nothing		
				end if
				Set objCurrency = Nothing
				%>
              <tr> 
			<form action="<%=Application("baseroot") & "/editor/currency/ListaCurrency.asp"%>" method="post" name="item_x_page">
			<th colspan="8" align="left">
			<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
			<%		
			'**************** richiamo paginazione
			call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/currency/ListaCurrency.asp", "&items="&itemsXpage)
			%>
			</th>
			</form>
              </tr>
		</table>
		<br/>	
		<form action="<%=Application("baseroot") & "/editor/currency/InserisciCurrency.asp"%>" method="post" name="form_crea">
		<input type="hidden" value="LCY" name="cssClass">	
		<input type="hidden" value="-1" name="id_currency">
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.currency.lista.button.label.inserisci")%>" onclick="javascript:document.form_crea.submit();" />	

		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.currency.lista.button.label.aggiorna")%>" onclick="javascript:location.href='<%=Application("baseroot") & "/editor/currency/refreshCurrency.asp"%>';" />
		</form>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>