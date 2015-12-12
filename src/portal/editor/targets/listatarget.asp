<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function deleteTarget(id_objref,row,refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.target.lista.js.alert.delete_target")%>?")){
		ajaxDeleteItem(id_objref,"target",row,refreshrows);
	}
}

function isNumerico(inputStr) {	
	for (var i = 0; i < inputStr.length; i++) {
		var oneChar = inputStr.substring(i, i + 1)
		if (oneChar < "0" || oneChar > "9") {
			return false;
		}
	}
	return true;
}

function isCharacterLowerCase(inputStr) {
	var oneChar = inputStr;
	if (oneChar < 97 || oneChar > 122) {
		return false;
	}
	return true;
}

//consente di digitare numeri e il punto
function isCorrectChar(e){
	var key = window.event ? e.keyCode : e.which;
	var keychar = String.fromCharCode(key);		
	if (isNumerico(keychar) || isCharacterLowerCase(key) || key==95 || keychar=="-"){					
		return true;
	}
	return false;
}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LT"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<div id="ajaxresp" align="center" style="background-color:#FFFF00; border:1px solid #000000; color:#000000; display:none;"></div>
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">
		      <tr> 
		      <th colspan="2">&nbsp;</th>
			  <th><%=langEditor.getTranslated("backend.target.lista.table.header.descrizione")%></th>
		      <th><%=langEditor.getTranslated("backend.target.lista.table.header.target_type")%></th>
		      <th><%=langEditor.getTranslated("backend.target.lista.table.header.automatic")%></th>
		      </tr> 
				<%
				Set objListaTarget = objTarget.getListaTarget()
				Dim intCount
				intCount = 0
				
				Dim newsCounter, iIndex, objTmpTarget, objTmpTargetKey, FromTarget, ToTarget, Diff
				iIndex = objListaTarget.Count
				FromTarget = ((numPage * itemsXpage) - itemsXpage)
				Diff = (iIndex - ((numPage * itemsXpage)-1))
				if(Diff < 1) then
					Diff = 1
				end if
				
				ToTarget = iIndex - Diff
				
				totPages = iIndex\itemsXpage
				if(totPages < 1) then
					totPages = 1
				elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
					totPages = totPages +1	
				end if		
				
				Dim styleRow, styleRow2
				styleRow2 = "table-list-on"
										
				objTmpTarget = objListaTarget.Items
				objTmpTargetKey=objListaTarget.Keys		
				for newsCounter = FromTarget to ToTarget
					styleRow = "table-list-off"
					if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>
				<form action="<%=Application("baseroot") & "/editor/targets/InserisciTarget.asp"%>" method="post" name="form_lista_<%=intCount%>">
				<input type="hidden" value="<%=objTmpTargetKey(newsCounter)%>" name="id_target">
				<input type="hidden" value="" name="delete_target">
				<input type="hidden" value="LT" name="cssClass">
				</form>	
				<tr class="<%=styleRow%>" id="tr_delete_list_<%=intCount%>">
					<%
					Set objTmpTarget0 = objTmpTarget(newsCounter)
					%>	
					<td align="center" width="25"><%if(objTmpTarget0.getTargetType() <> 3) then%><a href="javascript:document.form_lista_<%=intCount%>.submit();"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.target.lista.table.alt.modify_target")%>" hspace="2" vspace="0" border="0"></a><%end if%></td>
					<td align="center" width="25">
					<%if(objTmpTarget0.isLocked() = 0) then%>
						<a href="javascript:deleteTarget(<%=objTmpTargetKey(newsCounter)%>, 'tr_delete_list_<%=intCount%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.target.lista.table.alt.delete_target")%>" hspace="2" vspace="0" border="0"></a>
					<%else%>
						<img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.target.lista.table.alt.cant_delete_target")%>" hspace="2" vspace="0" border="0">				
					<%end if%></td>
					<td width="40%">
					<%if(objTmpTarget0.isLocked() = 0 AND not(objTmpTarget0.getTargetType()=3)) then%>					
					<div class="ajax" id="view_descrizione_<%=intCount%>" onmouseover="javascript:showHide('view_descrizione_<%=intCount%>','edit_descrizione_<%=intCount%>','descrizione_<%=intCount%>',500, false);"><%=objTmpTarget0.getTargetDescrizione()%></div>
					<div class="ajax" id="edit_descrizione_<%=intCount%>"><input type="text" class="formfieldAjax" id="descrizione_<%=intCount%>" name="descrizione" onmouseout="javascript:restoreField('edit_descrizione_<%=intCount%>','view_descrizione_<%=intCount%>','descrizione_<%=intCount%>','target',<%=objTmpTarget0.getTargetID()%>,1,<%=intCount%>);" value="<%=objTmpTarget0.getTargetDescrizione()%>" onkeypress="javascript:return isCorrectChar(event);"></div>
					<script>
					$("#edit_descrizione_<%=intCount%>").hide();
					</script>
					<%else
						response.write(objTmpTarget0.getTargetDescrizione())
					end if
					%>
					</td>
					<td>
					<%if(objTmpTarget0.isLocked() = 0 AND not(objTmpTarget0.getTargetType()=3)) then%>
					<div class="ajax" id="view_target_type_<%=intCount%>" onmouseover="javascript:showHide('view_target_type_<%=intCount%>','edit_target_type_<%=intCount%>','target_type_<%=intCount%>',500, true);">
					<%
					Select Case objTmpTarget0.getTargetType()
					Case 1
						response.write(langEditor.getTranslated("backend.target.lista.table.label.type_cat"))
'<!--nsys-trglist1-->
					Case 2
						response.write(langEditor.getTranslated("backend.target.lista.table.label.type_prod"))
'<!---nsys-trglist1-->
					Case 3
						response.write(langEditor.getTranslated("backend.target.lista.table.label.type_lang"))
					Case Else
					End Select%>
					</div>
					<div class="ajax" id="edit_target_type_<%=intCount%>">					
					  <%
					  Dim objTTmp
					  Set objTTmp = new TargetClass
					  Set typeTarget = objTTmp.getListaTargetType()%>
						<select name="target_type" class="formfieldAjaxSelect" id="target_type_<%=intCount%>" onblur="javascript:updateField('edit_target_type_<%=intCount%>','view_target_type_<%=intCount%>','target_type_<%=intCount%>','target',<%=objTmpTarget0.getTargetID()%>,2,<%=intCount%>);">
						<%if not (isNull(typeTarget)) then
							for each y in typeTarget.Keys
								if(y <> 3) then%>
								<option value="<%=y%>"<%if (y=objTmpTarget0.getTargetType()) then response.Write(" selected")%>><%=langEditor.getTranslated(typeTarget(y))%></option>	
							<%	end if
							next
						end if%>
						</SELECT>		  
					  <%Set typeTarget = nothing
					  Set objTTmp = nothing%>
					</div>
					<script>
					$("#edit_target_type_<%=intCount%>").hide();
					</script>
					<%else
						Select Case objTmpTarget0.getTargetType()
						Case 1
							response.write(langEditor.getTranslated("backend.target.lista.table.label.type_cat"))
'<!--nsys-trglist2-->
						Case 2
							response.write(langEditor.getTranslated("backend.target.lista.table.label.type_prod"))
'<!---nsys-trglist2-->
						Case 3
							response.write(langEditor.getTranslated("backend.target.lista.table.label.type_lang"))
						Case Else
						End Select
					end if
					%>
					</td>
					<td>
					<%if(objTmpTarget0.isLocked() = 0 AND not(objTmpTarget0.getTargetType()=3)) then%>
					<div class="ajax" id="view_automatic_<%=intCount%>" onmouseover="javascript:showHide('view_automatic_<%=intCount%>','edit_automatic_<%=intCount%>','automatic_<%=intCount%>',500, true);">
					<%
					Select Case objTmpTarget0.isAutomatic()
					Case 0
						response.write(langEditor.getTranslated("backend.commons.no"))
					Case 1
						response.write(langEditor.getTranslated("backend.commons.yes"))
					Case Else
					End Select%>
					</div>
					<div class="ajax" id="edit_automatic_<%=intCount%>">
						<select name="automatic" class="formFieldTXTShort" id="automatic_<%=intCount%>" onblur="javascript:updateField('edit_automatic_<%=intCount%>','view_automatic_<%=intCount%>','automatic_<%=intCount%>','target',<%=objTmpTarget0.getTargetID()%>,2,<%=intCount%>);">
						<option value="0"<%if ("0"=objTmpTarget0.isAutomatic()) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>	
						<option value="1"<%if ("1"=objTmpTarget0.isAutomatic()) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>
						</SELECT>						
					</div>
					<script>
					$("#edit_automatic_<%=intCount%>").hide();
					</script>
					<%else
						Select Case objTmpTarget0.isAutomatic()
						Case 0
							response.write(langEditor.getTranslated("backend.commons.no"))
						Case 1
							response.write(langEditor.getTranslated("backend.commons.yes"))
						Case Else
						End Select
					end if
					%>
					</td> 
				</tr>		
				<%intCount = intCount +1
				next
				Set objListaTarget = nothing
				Set objTarget = Nothing
				%>
			<tr> 
			<form action="<%=Application("baseroot") & "/editor/targets/ListaTarget.asp"%>" method="post" name="item_x_page">
			<th colspan="5">
			<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
			<%		
			'**************** richiamo paginazione
			call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/targets/ListaTarget.asp", "&items="&itemsXpage)
			%>
			</th>
			</form>
			</tr>
		</table>
		<br/>	
		<form action="<%=Application("baseroot") & "/editor/targets/InserisciTarget.asp"%>" method="post" name="form_crea">
		<input type="hidden" value="LT" name="cssClass">	
		<input type="hidden" value="-1" name="id_target">
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.target.lista.button.label.inserisci")%>" onclick="javascript:document.form_crea.submit();" />
		</form>		
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>