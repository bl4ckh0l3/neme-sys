<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<%
'<!--nsys-nwsletlist1-->
%>
<!-- #include virtual="/common/include/Objects/VoucherClass.asp" -->
<%
'<!---nsys-nwsletlist1-->
%>
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function deleteNewsletter(id_objref,row,refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.newsletter.lista.js.alert.delete_newsletter")%>?")){
		ajaxDeleteItem(id_objref,"newsletter",row,refreshrows);
	}
}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LNL"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<div id="ajaxresp" align="center" style="background-color:#FFFF00; border:1px solid #000000; color:#000000; display:none;"></div>
		<table border="0" align="top" cellpadding="0" cellspacing="0" class="principal">
              <tr> 
				  <th colspan="2">&nbsp;</th>
				  <th><%=langEditor.getTranslated("backend.newsletter.lista.table.header.descrizione")%></th>
				  <th><%=langEditor.getTranslated("backend.newsletter.lista.table.header.newsletter_stato")%></th>
				  <th><%=langEditor.getTranslated("backend.newsletter.lista.table.header.newsletter_template")%></th>
<!--nsys-nwsletlist2-->
				  <th><%=langEditor.getTranslated("backend.newsletter.lista.table.header.voucher_campaign")%></th>
<!---nsys-nwsletlist2-->
              </tr> 
				<%
				Dim totPages, hasNewsletter
				hasNewsletter = false
				on error Resume Next
				
					Set objListaNewsletter = objNewsletter.getListaNewsletter(null)
					
					if(objListaNewsletter.Count > 0) then
						hasNewsletter = true
					end if
					
				if Err.number <> 0 then
				end if	
				
				if(hasNewsletter) then
'<!--nsys-nwsletlist3-->				
				Set objVoucherClass =  new VoucherClass				
				On Error Resume Next
				hasVoucherCampaign = false
				Set objListVoucherCampaign = objVoucherClass.getCampaignList(4, 1)
				if(objListVoucherCampaign.count>0)then
					hasVoucherCampaign = true
				end if
				if(Err.number <> 0)then
					hasVoucherCampaign = false
				end if
				Set objVoucherClass = nothing
'<!---nsys-nwsletlist3-->
				
				Dim intCount
				intCount = 0
				
				Dim newsCounter, iIndex, objTmpNewsletter, objTmpNewsletterKey, FromNewsletter, ToNewsletter, Diff
				iIndex = objListaNewsletter.Count
				FromNewsletter = ((numPage * itemsXpage) - itemsXpage)
				Diff = (iIndex - ((numPage * itemsXpage)-1))
				if(Diff < 1) then
					Diff = 1
				end if
				
				ToNewsletter = iIndex - Diff
				
				totPages = iIndex\itemsXpage
				if(totPages < 1) then
					totPages = 1
				elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
					totPages = totPages +1	
				end if		

				Dim styleRow, styleRow2
				styleRow2 = "table-list-on"
						
				objTmpNewsletter = objListaNewsletter.Items
				objTmpNewsletterKey=objListaNewsletter.Keys		
				for newsCounter = FromNewsletter to ToNewsletter
					styleRow = "table-list-off"
					if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>
					<form action="<%=Application("baseroot") & "/editor/newsletter/InserisciNewsletter.asp"%>" method="post" name="form_lista_<%=intCount%>">
					<input type="hidden" value="<%=objTmpNewsletterKey(newsCounter)%>" name="id_newsletter">
					<input type="hidden" value="" name="delete_newsletter">
					<input type="hidden" value="LNL" name="cssClass">
					</form> 	
					<tr class="<%=styleRow%>" id="tr_delete_list_<%=intCount%>">
					<%
					Set objTmpNewsletter0 = objTmpNewsletter(newsCounter)
					%>	
					<td align="center" width="25"><a href="javascript:document.form_lista_<%=intCount%>.submit();"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.newsletter.lista.table.alt.modify_newsletter")%>" hspace="2" vspace="0" border="0"></a></td>
					<td align="center" width="25"><a href="javascript:deleteNewsletter(<%=objTmpNewsletterKey(newsCounter)%>, 'tr_delete_list_<%=intCount%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.newsletter.lista.table.alt.delete_newsletter")%>" hspace="2" vspace="0" border="0"></a></td>
					<td width="30%">
					<div class="ajax" id="view_desc_<%=intCount%>" onmouseover="javascript:showHide('view_desc_<%=intCount%>','edit_desc_<%=intCount%>','descrizione_<%=intCount%>',500, false);"><%=objTmpNewsletter0.getDescrizione()%></div>
					<div class="ajax" id="edit_desc_<%=intCount%>"><input type="text" class="formfieldAjax" id="descrizione_<%=intCount%>" name="descrizione" onmouseout="javascript:restoreField('edit_desc_<%=intCount%>','view_desc_<%=intCount%>','descrizione_<%=intCount%>','newsletter',<%=objTmpNewsletter0.getNewsletterID()%>,1,<%=intCount%>);" value="<%=objTmpNewsletter0.getDescrizione()%>"></div>
					<script>
					$("#edit_desc_<%=intCount%>").hide();
					</script>
					</td>
					<td width="10%">
					<div class="ajax" id="view_stato_<%=intCount%>" onmouseover="javascript:showHide('view_stato_<%=intCount%>','edit_stato_<%=intCount%>','stato_<%=intCount%>',500, true);">
					<%
					Select Case objTmpNewsletter0.getStato()
					Case 0
						response.write(langEditor.getTranslated("backend.newsletter.lista.table.label.inactive"))
					Case 1
						response.write(langEditor.getTranslated("backend.newsletter.lista.table.label.active"))
					Case Else
					End Select%>
					</div>
					<div class="ajax" id="edit_stato_<%=intCount%>">
					<select name="stato" class="formfieldAjaxSelect" id="stato_<%=intCount%>" onblur="javascript:updateField('edit_stato_<%=intCount%>','view_stato_<%=intCount%>','stato_<%=intCount%>','newsletter',<%=objTmpNewsletter0.getNewsletterID()%>,2,<%=intCount%>);">
					<option value="0"<%if (0=Cint(objTmpNewsletter0.getStato())) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.newsletter.lista.table.label.inactive")%></option>	
					<option value="1"<%if (1=Cint(objTmpNewsletter0.getStato())) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.newsletter.lista.table.label.active")%></option>	
					</SELECT>	
					</div>
					<script>
					$("#edit_stato_<%=intCount%>").hide();
					</script>
					</td>
					<td width="20%">
					<div class="ajax" id="view_template_<%=intCount%>" onmouseover="javascript:showHide('view_template_<%=intCount%>','edit_template_<%=intCount%>','template_<%=intCount%>',500, true);">
					<%=objTmpNewsletter0.getTemplate()%>
					</div>
					<div class="ajax" id="edit_template_<%=intCount%>">
					<select name="template" class="formfieldAjaxSelect" id="template_<%=intCount%>" onblur="javascript:updateField('edit_template_<%=intCount%>','view_template_<%=intCount%>','template_<%=intCount%>','newsletter',<%=objTmpNewsletter0.getNewsletterID()%>,2,<%=intCount%>);">		  
					<%
					dim listTemplate
					listTemplate = objNewsletter.getListaTemplateNewsletter()
					For y=LBound(listTemplate) to UBound(listTemplate)%>					
					<option value="<%=response.Write(listTemplate(y))%>"<%if (listTemplate(y)=objTmpNewsletter0.getTemplate()) then response.Write(" selected")%>><%=response.Write(listTemplate(y))%></option>	
					<%Next%>
					</SELECT>	
					</div>
					<script>
					$("#edit_template_<%=intCount%>").hide();
					</script>
				</td>
<!--nsys-nwsletlist4-->
					<td>
					<div class="ajax" id="view_voucher_<%=intCount%>" onmouseover="javascript:showHide('view_voucher_<%=intCount%>','edit_voucher_<%=intCount%>','voucher_<%=intCount%>',500, true);">
					<%=objListVoucherCampaign(objTmpNewsletter0.getVoucher()).getLabel()%>
					</div>
					<div class="ajax" id="edit_voucher_<%=intCount%>">
					<select name="voucher" class="formfieldAjaxSelect" id="voucher_<%=intCount%>" onblur="javascript:updateField('edit_voucher_<%=intCount%>','view_voucher_<%=intCount%>','voucher_<%=intCount%>','newsletter',<%=objTmpNewsletter0.getNewsletterID()%>,2,<%=intCount%>);">		  
					  <option value=""></option>
					  <%
					  if(hasVoucherCampaign)then
						for each g in objListVoucherCampaign%>
						<option value="<%=g%>" <%if(g=objTmpNewsletter0.getVoucher())then response.write(" selected") end if%>><%=objListVoucherCampaign(g).getLabel()%></option>
						<%next
					  end if
					  %>
					</SELECT>	
					</div>
					<script>
					$("#edit_voucher_<%=intCount%>").hide();
					</script>
				</td>               
<!---nsys-nwsletlist4-->
              	</tr>			
					<%intCount = intCount +1
				next
				Set objListaNewsletter = nothing
				end if
				Set objNewsletter = Nothing
				%>
              <tr> 
			<form action="<%=Application("baseroot") & "/editor/newsletter/ListaNewsletter.asp"%>" method="post" name="item_x_page">
<!--nsys-nwsletlist5-->
			<th colspan="6">
<!---nsys-nwsletlist5-->
			<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
			<%		
			'**************** richiamo paginazione
			call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/newsletter/ListaNewsletter.asp", "&items="&itemsXpage)
			%>
			</th>
			</form>
              </tr>
		</table>
		<br/>	
		<form action="<%=Application("baseroot") & "/editor/newsletter/InserisciNewsletter.asp"%>" method="post" name="form_crea">
		<input type="hidden" value="LNL" name="cssClass">	
		<input type="hidden" value="-1" name="id_newsletter">
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.newsletter.lista.button.label.inserisci")%>" onclick="javascript:document.form_crea.submit();" />
		</form>		
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>