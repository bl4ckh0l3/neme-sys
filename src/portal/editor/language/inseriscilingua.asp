<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/initlang.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function insertLanguage(){
	var strOptionValue = document.form_inserisci.option_code.options[document.form_inserisci.option_code.selectedIndex].value;
	document.form_inserisci.descrizione.value = strOptionValue.substring(0, strOptionValue.indexOf("|"));
	document.form_inserisci.selected_label.value = strOptionValue.substring(strOptionValue.indexOf("|")+1,strOptionValue.length);
	
	if(confirm("<%=langEditor.getTranslated("backend.lingue.js.alert.confirm_set_target_to_user")%>")){
		document.form_inserisci.set_target_to_users.value = 1
	}
	
	var element = document.getElementById("urlSubdomain");
	if(element.style.visibility == 'visible'){
		if(document.form_inserisci.url_subdomain.value == ''){
			alert("<%=langEditor.getTranslated("backend.lingue.js.alert.empty_url")%>");
			return;
		}		
	}	
	
	document.form_inserisci.submit();
}

function deleteLanguage(id_language,description){
/*<!--nsys-demolangtmp1-->*/
	document.form_delete_lang.id_language.value = id_language;
	document.form_delete_lang.descrizione.value = description;
	document.form_delete_lang.delete_language.value = "del";
	document.form_delete_lang.action = "<%=Application("baseroot") & "/editor/language/ProcessLanguage.asp"%>";
	document.form_delete_lang.submit();
/*<!---nsys-demolangtmp1-->*/
}

function activateLanguage(id_language,elem){
/*<!---nsys-demolangtmp2-->*/
	document.form_activate_lang.id_lang_to_activate.value = id_language;
	document.form_activate_lang.lang_to_active.value = elem.value;
	document.form_activate_lang.submit();
/*<!---nsys-demolangtmp2-->*/
}

function showHideURLsubdomain(elemID){
	var elem = document.form_inserisci.subdomain_active.options[document.form_inserisci.subdomain_active.selectedIndex].value;

	var element = document.getElementById(elemID);
	if(elem == 0){
		element.style.visibility = 'hidden';
		element.style.display = "none";
	}else if(elem == 1){
		element.style.visibility = 'visible';		
		element.style.display = "block";
	}
}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="IL"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table class="principal" border="0" cellpadding="0" cellspacing="0">
		      <tr> 
				<th>&nbsp;</td>
				<th><%=langEditor.getTranslated("backend.lingue.lista.table.header.descrizione")%></th>
				<th><%=langEditor.getTranslated("backend.lingue.lista.table.header.lang_active")%></th>
				<th><%=langEditor.getTranslated("backend.lingue.lista.table.header.subdomain_active")%></th>
				<th><%=langEditor.getTranslated("backend.lingue.lista.table.header.url_sottodominio")%></th>
		      </tr> 
				<%
				On Error Resume Next
				Set objListaLanguage = objLanguage.getListaLanguage()					
				if isObject(objListaLanguage) AND not(isEmpty(objListaLanguage)) then
					Dim intCount
					intCount = 0
					
					Dim langCounter, iIndex, objTmpLanguageItem, FromLanguage, ToLanguage, Diff
					iIndex = objListaLanguage.Count
					FromLanguage = ((numPage * itemsXpage) - itemsXpage)
					Diff = (iIndex - ((numPage * itemsXpage)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToLanguage = iIndex - Diff
					
					totPages = iIndex\itemsXpage
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
						totPages = totPages +1	
					end if		
				
					Dim styleRow, styleRow2
					styleRow2 = "table-list-on"
							
					objTmpLanguageItem=objListaLanguage.Items	
					for langCounter = FromLanguage to ToLanguage
						styleRow = "table-list-off"
						if(langCounter MOD 2 = 0) then styleRow = styleRow2 end if
						Set objThisLanguage = objTmpLanguageItem(langCounter)%>
						<tr class="<%=styleRow%>">
						<td align="center" width="25"><a href="javascript:deleteLanguage(<%=objThisLanguage.getLanguageID()%>,'<%=objThisLanguage.getLanguageDescrizione()%>');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.lingue.lista.table.alt.delete_lang")%>" hspace="2" vspace="0" border="0"></a></td>
						<td><img width="16" height="11" border="0" style="padding-right:5px;vertical-align:middle;" alt="<%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&objThisLanguage.getLabelDescrizione())%>" title="<%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&objThisLanguage.getLabelDescrizione())%>" src="/editor/img/flag/flag-<%=objThisLanguage.getLanguageDescrizione()%>.png"><%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&objThisLanguage.getLabelDescrizione())%></td>
						<td>
						  <select name="lang_to_active" class="formFieldTXTShort" onChange="activateLanguage(<%=objThisLanguage.getLanguageID()%>,this);">
						  <option value="0" <%if (objThisLanguage.isLangActive() = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>
						  <option value="1" <%if (objThisLanguage.isLangActive() = 1) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>
						  </select>
						</td>

						<td><%if (objThisLanguage.isSubDomainActive()) then response.Write(langEditor.getTranslated("backend.commons.yes")) else response.Write(langEditor.getTranslated("backend.commons.no"))%></td>
						<td><%=objThisLanguage.getURLSubdomain()%></td>
						</tr>				
						<%intCount = intCount +1
						Set objThisLanguage = nothing
					next
					
				end if
				
				if(Err.number <> 0) then
				
				end if
				%>
			<form action="<%=Application("baseroot") & "/editor/language/InserisciLingua.asp"%>" method="post" name="item_x_page">
              <tr> 
                <th colspan="5" align="left">
				<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
				<%		
				'**************** richiamo paginazione
				call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/language/InserisciLingua.asp", "&items="&itemsXpage)
				%>
				</th>
              </tr>
			</form>
		</table>
			<br/><br>
			<%Dim objLangDisponibili
			if(Instr(1, typename(objLanguage.getListaLangDisponibili()), "dictionary", 1) > 0) then
			Set objLangDisponibili = objLanguage.getListaLangDisponibili()%> 
			<table border="0" align="top" cellpadding="0" cellspacing="0" class="principal">
			<form action="<%=Application("baseroot") & "/editor/language/ProcessLanguage.asp"%>" method="post" name="form_inserisci">
			<input type="hidden" value="-1" name="id_language">
			<input type="hidden" value="" name="descrizione">
			<input type="hidden" value="" name="selected_label">
			<input type="hidden" value="0" name="set_target_to_users">		
			<tr>		  
				<td align="left" valign="top">
				 <span class="labelForm"><%=langEditor.getTranslated("backend.lingue.lista.label.lang_name")%></span><br>
				 <select name="option_code" class="formFieldTXT">
					<%for each x in objLangDisponibili%>	
						<%if not(objLanguage.isLanguageSelected(x)) then%><option value="<%=x&"|"&objLangDisponibili(x)%>" style="background-image: url(<%=Application("baseroot") & "/editor/img/flag/flag-"&x&".png"%>);background-repeat: no-repeat;background-position: left center;padding-left:20px;padding-bottom:2px;vertical-align:top;"><%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&objLangDisponibili(x))%></option><%end if%>
					<%next%>
					</select>&nbsp;&nbsp;
				</td>
				<td align="left" valign="top">&nbsp;</td>
			</tr>
			<tr>  
				<td align="left" valign="top">
				<br><span class="labelForm"><%=langEditor.getTranslated("backend.lingue.lista.label.lang_active")%></span><br>
			      <select name="lang_active" class="formFieldTXTShort">
				  <option value="0"><%=langEditor.getTranslated("backend.commons.no")%></option>
				  <option value="1"><%=langEditor.getTranslated("backend.commons.yes")%></option>
				</select>&nbsp;&nbsp;
				</td>
				<td align="left" valign="top">&nbsp;</td>
			</tr>
			<tr>  
				<td align="left" valign="top">
				<br><span class="labelForm"><%=langEditor.getTranslated("backend.lingue.lista.label.subdomain_active")%></span><br>
			      <select name="subdomain_active" class="formFieldTXTShort" onChange="javascript:showHideURLsubdomain('urlSubdomain')">
				  <option value="0"><%=langEditor.getTranslated("backend.commons.no")%></option>
				  <option value="1"><%=langEditor.getTranslated("backend.commons.yes")%></option>
				</select>&nbsp;&nbsp;
				</td>
				<td align="left" valign="top">
				<br><div id="urlSubdomain" style="visibility:hidden;display:none;" align="left"> 
				<span class="labelForm"><%=langEditor.getTranslated("backend.lingue.lista.label.url_subdomain_active")%></span><br>
				<input type="text" name="url_subdomain" class="formFieldTXTLong" value="">
				</div> 
			  	</td>
			</tr>
			</form>
			</table>
			<br/>
			<input type="button" class="buttonForm" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.lingue.lista.button.label.inserisci")%>" onclick="javascript:insertLanguage();" />
			<br/><br/>

			<form action="" method="post" name="form_delete_lang">
			<input type="hidden" value="" name="id_language">
			<input type="hidden" value="" name="delete_language">
			<input type="hidden" value="" name="descrizione">
			</form>
			
			<form action="<%=Application("baseroot") & "/editor/language/ActivateLang.asp"%>" method="post" name="form_activate_lang">
			<input type="hidden" value="" name="id_lang_to_activate">
			<input type="hidden" value="<%=itemsXpage%>" name="items">			
			<input type="hidden" value="<%=numPage%>" name="page">		
			<input type="hidden" value="" name="lang_to_active">
			</form>

			<%
			Set objLangDisponibili = nothing
			end if
			Set objLanguage = Nothing%>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>