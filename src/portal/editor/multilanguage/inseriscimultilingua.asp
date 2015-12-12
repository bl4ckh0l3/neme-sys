<%@Language=VBScript codepage=65001 %>
<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/initmultilang.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function insertLanguage(){
	if(document.form_inserisci.keyword.value == ""){
		alert("<%=langEditor.getTranslated("backend.multilingue.lista.js.alert.insert_keyword")%>");
		return;
	}

	//verifico se si tratta di un messaggio javascript o di una label testuale
	var isJs = false;
	if(document.form_inserisci.keyword.value.indexOf(".js.") > 0){
		isJs = true;
	}
	
	var strTmpValue
	<%if (bolFoundLista) then
		for each k in objListaLanguage%>
	strTmpValue = document.form_inserisci.value_<%=k%>.value;
	strTmpValue = replaceChars(strTmpValue,isJs);
	document.form_inserisci.value_<%=k%>.value = strTmpValue;
	<%	next
	end if%>
	
	document.form_inserisci.submit();
}

function modifyLanguage(theForm){
	if(theForm.keyword.value == ""){
		alert("<%=langEditor.getTranslated("backend.multilingue.lista.js.alert.insert_keyword")%>");
		return;
	}

	//verifico se si tratta di un messaggio javascript o di una label testuale
	var isJs = false;
	if(theForm.keyword.value.indexOf(".js.") > 0){
		isJs = true;
	}
	
	var strTmpValue
	<%if (bolFoundLista) then
		for each k in objListaLanguage%>
	strTmpValue = document.form_inserisci.value_<%=k%>.value;
	strTmpValue = replaceChars(strTmpValue,isJs);
	document.form_inserisci.value_<%=k%>.value = strTmpValue;
	<%	next
	end if%>	
	
	theForm.submit();
}

function deleteLanguage(theForm){	
	if(confirm("<%=langEditor.getTranslated("backend.multilingue.lista.js.alert.confirm_delete_multilang")%>?")){
		theForm.operation.value = "delete";
		theForm.action = "<%=Application("baseroot") & "/editor/multilanguage/ProcessMultiLanguage.asp"%>";
		theForm.submit();
	}
}

function selectAllLanguageKey(){
	var form, ck_value, is_ck_value_ck;
	ck_value = document.getElementById("ck_do_select_all");
	is_ck_value_ck = ck_value.checked;
	for (var counter = 0; counter < <%=(itemsXpage)%>; counter++) {
		form = document.getElementById("form_lista_"+counter);
		if(form){			
			if (is_ck_value_ck){
				form.ck_select_all.checked = true;
			}else{
				form.ck_select_all.checked = false;
			}
		}
	}
}

function modifyAllSelectedLanguage(){
	var form;
	for (var counter = 0; counter < <%=(itemsXpage)%>; counter++) {
		form = document.getElementById("form_lista_"+counter);
		if(form){			
			if (form.ck_select_all.checked){
				var singleLineValue;
				singleLineValue = "id=" + form.id_multi_language.value + "||";
				singleLineValue += "keyword=" + form.keyword.value + "||";

				//verifico se si tratta di un messaggio javascript o di una label testuale
				var isJs = false;
				if(form.keyword.value.indexOf(".js.") > 0){
					isJs = true;
				}
				
				var strTmpValue;
				<%if (bolFoundLista) then
					for each k in objListaLanguage%>
						strTmpValue = "value_<%=k%>=" + form.value_<%=k%>.value;
						strTmpValue = replaceChars(strTmpValue,isJs);
						singleLineValue += strTmpValue + "||";				
				<%	next
				end if%>		
				
				singleLineValue = singleLineValue.substring(0,singleLineValue.lastIndexOf("||"));
				singleLineValue += "###";
				
				document.form_lista_multi_select.multiple_values.value += singleLineValue;
			}
		}
	}
	
	document.form_lista_multi_select.multiple_values.value = document.form_lista_multi_select.multiple_values.value.substring(0,document.form_lista_multi_select.multiple_values.value.lastIndexOf("###"));
	document.form_lista_multi_select.operation.value = "modify";

	if(confirm("<%=langEditor.getTranslated("backend.multilingue.lista.js.alert.confirm_modify_sel_multilang")%>?")){
		document.form_lista_multi_select.submit();
	}
}

function deleteAllSelectedLanguage(){
	var form;
	var singleLineValue = "";
	for (var counter = 0; counter < <%=(itemsXpage)%>; counter++) {
		form = document.getElementById("form_lista_"+counter);
		if(form){			
			if (form.ck_select_all.checked){				
				singleLineValue += form.id_multi_language.value + "|";
			}
		}
	}
				
	singleLineValue = singleLineValue.substring(0,singleLineValue.lastIndexOf("|"));			

	if(confirm("<%=langEditor.getTranslated("backend.multilingue.lista.js.alert.confirm_delete_sel_multilang")%>?")){
		document.form_lista_multi_select.multiple_values.value += singleLineValue;	
		document.form_lista_multi_select.operation.value = "delete";
		document.form_lista_multi_select.submit();	
	}
}

function replaceChars(inString,isJs){
	var outString = inString;
	var pos= 0;
	
	// ricerca e escaping degli apici	
	var quote= -1;
	do {
		quote= outString.indexOf('\'', pos);
		if (quote >= 0) {
			if(isJs){
				outString= outString.substring(0, quote) + "\'" + outString.substring(quote +1);
			}else{
				outString= outString.substring(0, quote) + "&#39;" + outString.substring(quote +1);
			}
			pos= quote+2;
		}
	} while (quote >= 0);

	// ricerca e escaping dei doppi apici
	pos= 0;
	var double_quote= -1;
	do {
		double_quote= outString.indexOf('"', pos);
		if (double_quote >= 0) {
			outString= outString.substring(0, double_quote) + "&quot;" + outString.substring(double_quote +1);
			pos= double_quote+2;
		}
	} while (double_quote >= 0);
	
	// ricerca e escaping dei new line
	pos= 0;
	var linefeed= -1;
	do {
		linefeed= outString.indexOf('\n', pos);
		if (linefeed >= 0) {
			outString= outString.substring(0, linefeed) + "\\n" + outString.substring(linefeed +1);
			pos= linefeed+2;
		}
	} while (linefeed >= 0);

	// ricerca e escaping dei line feed
	pos= 0;
	var creturn= -1;
	do {
		creturn= outString.indexOf('\r', pos);
		if (creturn >= 0) {
			outString= outString.substring(0, creturn) + "\\r" + outString.substring(creturn +1);
			pos= creturn+2;
		}
	} while (creturn >= 0);
	
	return outString;
}
</script>
</head>
<body onLoad="javascript:document.form_search.search_key.focus();">
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="IML"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table class="principal" border="0" cellpadding="0" cellspacing="0">
			<%if (bolFoundLista) then%>
				<form action="<%=Application("baseroot") & "/editor/multilanguage/inserisciMultilingua.asp"%>" method="post" name="form_search" accept-charset="UTF-8">
				<input type="hidden" value="1" name="page">
				  <tr height="35">
					<td colspan="2" align="center"><input type="submit" value="<%=langEditor.getTranslated("backend.multilingue.lista.label.search")%>" class="buttonForm" hspace="4"></td>
					<td colspan="<%=objListaLanguage.Count%>">
						<input type="text" name="search_key" value="<%=search_key%>" class="formFieldTXTLangKeyword">
					</td>			
				  </tr>
				</form>
				<tr> 
				<th colspan="2">&nbsp;</th>
				<th><%=langEditor.getTranslated("backend.multilingue.lista.table.header.keyword")%></th>
				<%for each k in objListaLanguage%>
					<th class="upper"><%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&objListaLanguage(k))%></th>
				<%next%>
				</tr> 
				<form action="<%=Application("baseroot") & "/editor/multilanguage/ProcessMultiLanguage.asp"%>" method="post" name="form_inserisci" accept-charset="UTF-8">
				<input type="hidden" value="-1" name="id_multi_language">
				<input type="hidden" value="" name="operation">
				<input type="hidden" value="0" name="is_multiple_selection">
				<input type="hidden" value="<%=itemsXpage%>" name="items">
				<input type="hidden" name="search_key" value="<%=search_key%>">
				<input type="hidden" value="<%=numPage%>" name="page">
				  <tr height="35">
					<td colspan="2" align="center"><input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.multilingue.lista.button.label.inserisci")%>" onclick="javascript:insertLanguage();" /></td>
					<td><input type="text" name="keyword" value="" class="formFieldTXTLangKeyword"></td>
					<%for each k in objListaLanguage%>
						<td><input type="text" name="value_<%=k%>" value="" class="formFieldTXTLang"></td>
					<%next%>
				  </tr>	
				</form>				
				<tr> 
				<th class="icons" width="70" nowrap>
					<a href="javascript:modifyAllSelectedLanguage();"><img src="<%=Application("baseroot")&"/editor/img/accept.png"%>" alt="<%=langEditor.getTranslated("backend.multilingue.lista.table.alt.modify_sel_lang")%>" hspace="2" vspace="0" border="0"></a>
					<input type="checkbox" value="" id="ck_do_select_all" name="ck_do_select_all" onclick="javascript:selectAllLanguageKey();"/>
					<a href="javascript:deleteAllSelectedLanguage();"><img src="<%=Application("baseroot")&"/editor/img/delete.png"%>" alt="<%=langEditor.getTranslated("backend.multilingue.lista.table.alt.delete_sel_lang")%>" hspace="2" vspace="0" border="0"></a>
				</th>
				<th width="70">&nbsp;</th>
				<th><%=langEditor.getTranslated("backend.multilingue.lista.table.header.keyword")%></th>
				<%for each k in objListaLanguage%>
					<th class="upper"><%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&objListaLanguage(k))%></th>
				<%next%>
				</tr> 
				<%

				Dim objListModifyKeys, hasList, totPages				
				hasList = false	

				on error resume next
				Set objListModifyKeys = objLanguage.searchDistinctKeyList(search_key)
				if(objListModifyKeys.Count > 0) then
					hasList = true
				end if

				if Err.number <> 0 then
					'response.write(Err.description)
				end if

				if(hasList) then									
					Dim intCount
					intCount = 0

					Dim listCounter, iIndex, objTmpList, objTmpListKey, FromList, ToList, Diff, objFilteredList
					iIndex = objListModifyKeys.Count
					FromList = ((numPage * itemsXpage) - itemsXpage)
					Diff = (iIndex - ((numPage * itemsXpage)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToList = iIndex - Diff
					
					totPages = iIndex\itemsXpage
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
						totPages = totPages +1	
					end if		
							
					objTmpListKeys = objListModifyKeys.Keys
					objTmpListItems = objListModifyKeys.Items		
						
					for listCounter = FromList to ToList					
						tmpID = objTmpListItems(listCounter)				
						tmpKey = objTmpListKeys(listCounter)
						Set filteredValuesList = objLanguage.searchFilteredListElementsByKey(tmpKey)%>
						<form action="<%=Application("baseroot") & "/editor/multilanguage/ProcessMultiLanguage.asp"%>" method="post" id="form_lista_<%=intCount%>" name="form_lista_<%=intCount%>" accept-charset="UTF-8">
						<tr>
						<td class="icons" width="70"><input type="checkbox" value="" name="ck_select_all"/></td>
						<td class="icons" width="70" nowrap>
						<input type="hidden" value="<%=tmpID%>" name="id_multi_language">
						<input type="hidden" value="" name="operation">
						<input type="hidden" value="0" name="is_multiple_selection">
						<input type="hidden" name="search_key" value="<%=search_key%>">
						<input type="hidden" value="<%=itemsXpage%>" name="items">
						<input type="hidden" value="<%=numPage%>" name="page"> 
						<a href="javascript:modifyLanguage(document.form_lista_<%=intCount%>);"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.multilingue.lista.table.alt.modify_lang")%>" hspace="5" vspace="0" border="0"></a><a href="javascript:deleteLanguage(document.form_lista_<%=intCount%>);"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.multilingue.lista.table.alt.delete_lang")%>" hspace="5" vspace="0" border="0"></a>
						</td>
						<td><input type="text" name="keyword" value="<%=tmpKey%>" class="formFieldTXTLangKeyword"></td>
						<%for each k in objListaLanguage%>
						<td><input type="text" name="value_<%=k%>" value="<%=filteredValuesList.item(k&"-"&tmpKey)%>" class="formFieldTXTLang"></td>	
						<%next%>
						</tr>				
						</form>						
						<%intCount = intCount +1	
						Set filteredValuesList = nothing
						Set objFilteredList = nothing
					Next

				end if%> 


				<form action="<%=Application("baseroot") & "/editor/multilanguage/InserisciMultilingua.asp"%>" method="post" name="item_x_page">
				<input type="hidden" name="search_key" value="<%=search_key%>">
			      <tr> 
				<th colspan="<%=3+objListaLanguage.Count%>" align="left">
				<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
				<%		
				'**************** richiamo paginazione
				call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/multilanguage/InserisciMultiLingua.asp", "&search_key="&search_key&"&items="&itemsXpage)
				%>
				</th>
			      </tr>
				</form>
				
				<%		 
				if Err.number <> 0 then
					response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
				end if
				
				Set objListaLanguage = nothing				
			end if
			Set objLanguage = Nothing%>
		</table>
		<br/><br/>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>

<form action="<%=Application("baseroot") & "/editor/multilanguage/ProcessMultiLanguage.asp"%>" method="post" name="form_lista_multi_select" accept-charset="UTF-8">
<input type="hidden" value="1" name="is_multiple_selection">
<input type="hidden" value="" name="operation">
<input type="hidden" value="" name="multiple_values">
<input type="hidden" value="<%=itemsXpage%>" name="items">
<input type="hidden" name="search_key" value="<%=search_key%>">
</form>
</body>
</html>