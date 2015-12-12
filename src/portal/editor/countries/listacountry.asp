<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function deleteCountry(id_objref,row,refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.country.lista.js.alert.delete_country")%>?")){
		ajaxDeleteItem(id_objref,"country",row,refreshrows);
	}
}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LCT"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<div id="ajaxresp" align="center" style="background-color:#FFFF00; border:1px solid #000000; color:#000000; display:none;"></div>
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">
			<form action="<%=Application("baseroot") & "/editor/countries/listacountry.asp"%>" method="post" name="form_search">
			<input type="hidden" value="1" name="page">
			  <tr height="35">
				<td colspan="8">
					<input type="submit" maxlength="100" value="<%=langEditor.getTranslated("backend.country.lista.label.search")%>" class="buttonForm" hspace="4" align="absbottom">&nbsp;&nbsp;<input type="text" name="search_key" value="<%=search_key%>" class="formFieldTXTLangKeyword">
				</td>			
			  </tr>
			</form>
			<tr> 
				  <th colspan="2">&nbsp;</th>
				  <th><%=langEditor.getTranslated("backend.country.lista.table.header.country_code")%></th>
				  <th><%=langEditor.getTranslated("backend.country.lista.table.header.country")%></th>
				  <th><%=langEditor.getTranslated("backend.country.lista.table.header.state_region_code")%></th>
				  <th><%=langEditor.getTranslated("backend.country.lista.table.header.state_region")%></th>
				  <th><%=langEditor.getTranslated("backend.country.lista.table.header.active")%></th>
				  <th><%=langEditor.getTranslated("backend.country.lista.table.header.use_for")%></th>
			</tr> 
				<%
				On Error Resume Next
				Dim hasCountry
				hasCountry = false

				if(search_key<>"")then
					Set objListaCountry = objCountry.findCountry(search_key)	
					hasCountry = true	
				else
					Set objListaCountry = objCountry.getListaCountry(null,null,null,null)				
					hasCountry = true	
				end if			
				
				if Err.number <> 0 then
					hasCountry = false
				end if
				
				if(hasCountry) then			
					Dim intCount
					intCount = 0
					
					Dim newsCounter, iIndex, objTmpCountry, objTmpCountryKey, FromCountry, ToCountry, Diff
					iIndex = objListaCountry.Count
					FromCountry = ((numPage * itemsXpage) - itemsXpage)
					Diff = (iIndex - ((numPage * itemsXpage)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToCountry = iIndex - Diff
					
					totPages = iIndex\itemsXpage
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
						totPages = totPages +1	
					end if		
				
					Dim styleRow, styleRow2
					styleRow2 = "table-list-on"							
							
					objTmpCountry = objListaCountry.Items
					objTmpCountryKey=objListaCountry.Keys	
					for newsCounter = FromCountry to ToCountry
						styleRow = "table-list-off"
						if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>
						<form action="<%=Application("baseroot") & "/editor/countries/InserisciCountry.asp"%>" method="post" name="form_lista_<%=intCount%>">
						<input type="hidden" value="<%=objTmpCountryKey(newsCounter)%>" name="id_country">
						<input type="hidden" value="" name="delete_country">
						<input type="hidden" value="LCT" name="cssClass">
						</form> 
					<tr class="<%=styleRow%>" id="tr_delete_list_<%=intCount%>">
						<%Set objTmpCountry0 = objTmpCountry(newsCounter)%>	
						<td align="center" width="25"><a href="javascript:document.form_lista_<%=intCount%>.submit();"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.country.lista.table.alt.modify_country")%>" hspace="2" vspace="0" border="0"></a></td>
						<td align="center" width="25"><a href="javascript:deleteCountry(<%=objTmpCountryKey(newsCounter)%>, 'tr_delete_list_<%=intCount%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.country.lista.table.alt.delete_country")%>" hspace="2" vspace="0" border="0"></a></td>
						<td width="10%">						
						<div class="ajax" id="view_country_code_<%=intCount%>" onmouseover="javascript:showHide('view_country_code_<%=intCount%>','edit_country_code_<%=intCount%>','country_code_<%=intCount%>',500, false);"><%=objTmpCountry0.getCountryCode()%></div>
						<div class="ajax" id="edit_country_code_<%=intCount%>"><input type="text" class="formfieldAjaxShort" id="country_code_<%=intCount%>" name="country_code" onmouseout="javascript:restoreField('edit_country_code_<%=intCount%>','view_country_code_<%=intCount%>','country_code_<%=intCount%>','country',<%=objTmpCountry0.getID()%>,1,<%=intCount%>);" value="<%=objTmpCountry0.getCountryCode()%>"></div>
						<script>
						$("#edit_country_code_<%=intCount%>").hide();
						</script>
						</td>
						<td width="19%">						
						<div class="ajax" id="view_country_description_<%=intCount%>" onmouseover="javascript:showHide('view_country_description_<%=intCount%>','edit_country_description_<%=intCount%>','country_description_<%=intCount%>',500, false);"><%=Server.HTMLEncode(objTmpCountry0.getCountryDescription())%></div>
						<div class="ajax" id="edit_country_description_<%=intCount%>"><input type="text" class="formfieldAjax" id="country_description_<%=intCount%>" name="country_description" onmouseout="javascript:restoreField('edit_country_description_<%=intCount%>','view_country_description_<%=intCount%>','country_description_<%=intCount%>','country',<%=objTmpCountry0.getID()%>,1,<%=intCount%>);" value="<%=objTmpCountry0.getCountryDescription()%>"></div>
						<script>
						$("#edit_country_description_<%=intCount%>").hide();
						</script>
						</td>
						
						<td width="16%">						
						<div class="ajax" id="view_state_region_code_<%=intCount%>" onmouseover="javascript:showHide('view_state_region_code_<%=intCount%>','edit_state_region_code_<%=intCount%>','state_region_code_<%=intCount%>',500, false);"><%=objTmpCountry0.getStateRegionCode()%></div>
						<div class="ajax" id="edit_state_region_code_<%=intCount%>"><input type="text" class="formfieldAjax" id="state_region_code_<%=intCount%>" name="state_region_code" onmouseout="javascript:restoreField('edit_state_region_code_<%=intCount%>','view_state_region_code_<%=intCount%>','state_region_code_<%=intCount%>','country',<%=objTmpCountry0.getID()%>,1,<%=intCount%>);" value="<%=objTmpCountry0.getStateRegionCode()%>"></div>
						<script>
						$("#edit_state_region_code_<%=intCount%>").hide();
						</script>
						</td>
						<td width="25%">						
						<div class="ajax" id="view_state_region_description_<%=intCount%>" onmouseover="javascript:showHide('view_state_region_description_<%=intCount%>','edit_state_region_description_<%=intCount%>','state_region_description_<%=intCount%>',500, false);"><%=Server.HTMLEncode(objTmpCountry0.getStateRegionDescription())%></div>
						<div class="ajax" id="edit_state_region_description_<%=intCount%>"><input type="text" class="formfieldAjaxLong" id="state_region_description_<%=intCount%>" name="state_region_description" onmouseout="javascript:restoreField('edit_state_region_description_<%=intCount%>','view_state_region_description_<%=intCount%>','state_region_description_<%=intCount%>','country',<%=objTmpCountry0.getID()%>,1,<%=intCount%>);" value="<%=objTmpCountry0.getStateRegionDescription()%>"></div>
						<script>
						$("#edit_state_region_description_<%=intCount%>").hide();
						</script>
						</td>						
						
						<td width="5%">
						<div class="ajax" id="view_active_<%=intCount%>" onmouseover="javascript:showHide('view_active_<%=intCount%>','edit_active_<%=intCount%>','active_<%=intCount%>',500, true);">
						<%
						if (strComp("1", objTmpCountry0.isActive(), 1) = 0) then 
							response.Write(langEditor.getTranslated("backend.commons.yes"))
						else 
							response.Write(langEditor.getTranslated("backend.commons.no"))
						end if
						%>
						</div>
						<div class="ajax" id="edit_active_<%=intCount%>">
						<select name="active" class="formfieldAjaxSelect" id="active_<%=intCount%>" onblur="javascript:updateField('edit_active_<%=intCount%>','view_active_<%=intCount%>','active_<%=intCount%>','country',<%=objTmpCountry0.getID()%>,2,<%=intCount%>);">
						<OPTION VALUE="0" <%if (strComp("0", objTmpCountry0.isActive(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
						<OPTION VALUE="1" <%if (strComp("1", objTmpCountry0.isActive(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
						</SELECT>	
						</div>
						<script>
						$("#edit_active_<%=intCount%>").hide();
						</script>
						</td>
						<td>
						<div class="ajax" id="view_use_for_<%=intCount%>" onmouseover="javascript:showHide('view_use_for_<%=intCount%>','edit_use_for_<%=intCount%>','use_for_<%=intCount%>',500, true);">
						<%
						Select case objTmpCountry0.getUseFor()
							Case 1
							response.Write(langEditor.getTranslated("backend.country.use_for.registration"))
							Case 2
							response.Write(langEditor.getTranslated("backend.country.use_for.purchase"))
							Case 3
							response.Write(langEditor.getTranslated("backend.country.use_for.all"))
							Case Else
						End Select
						%>
						</div>
						<div class="ajax" id="edit_use_for_<%=intCount%>">
						<select name="use_for" class="formfieldAjaxSelect" id="use_for_<%=intCount%>" onblur="javascript:updateField('edit_use_for_<%=intCount%>','view_use_for_<%=intCount%>','use_for_<%=intCount%>','country',<%=objTmpCountry0.getID()%>,2,<%=intCount%>);">
						<option value="1"<%if ("1"=objTmpCountry0.getUseFor()) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.country.use_for.registration")%></option>	
<!--nsys-cntlist1-->
						<option value="2"<%if ("2"=objTmpCountry0.getUseFor()) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.country.use_for.purchase")%></option>	
						<option value="3"<%if ("3"=objTmpCountry0.getUseFor()) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.country.use_for.all")%></option>	
<!---nsys-cntlist1-->
						</SELECT>	
						</div>
						<script>
						$("#edit_use_for_<%=intCount%>").hide();
						</script>
						</td>               
						</tr>				
						<%intCount = intCount +1
						
						Set objTmpCountry0 = nothing
					next
					Set objListaCountry = nothing		
				end if
				Set objCountry = Nothing
				%>
		<tr> 
			<form action="<%=Application("baseroot") & "/editor/countries/ListaCountry.asp"%>" method="post" name="item_x_page">
			<th colspan="8">
			<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
			<%		
			'**************** richiamo paginazione
			call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/countries/ListaCountry.asp", "&items="&itemsXpage)
			%>
			</th>
			</form>
              </tr>
		</table>
		<br/>	
		<form action="<%=Application("baseroot") & "/editor/countries/InserisciCountry.asp"%>" method="post" name="form_crea">
		<input type="hidden" value="LCT" name="cssClass">	
		<input type="hidden" value="-1" name="id_country">
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.country.lista.button.label.inserisci")%>" onclick="javascript:document.form_crea.submit();" />
		</form>		
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>