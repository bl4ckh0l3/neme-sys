<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include file="include/init2.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function insertCategoria(){
	if(controllaCampiInput()){
		document.form_inserisci.submit();
	}else{
		return;
	}
}

function move(fbox, tbox){
	var arrFbox = new Array();
	var arrTbox = new Array();
	var arrLookup = new Array();
	var i;
	
	for(i = 0; i < tbox.options.length; i++){
		arrLookup[tbox.options[i].text] = tbox.options[i].value;
		arrTbox[i] = tbox.options[i].text;
	}
	
	var fLength = 0;
	var tLength = arrTbox.length;
	
	for(i = 0; i < fbox.options.length; i++){
		arrLookup[fbox.options[i].text] = fbox.options[i].value;
		if(fbox.options[i].selected && fbox.options[i].value != ""){
			arrTbox[tLength] = fbox.options[i].text;
			tLength++;
		}else{
			arrFbox[fLength] = fbox.options[i].text;
			fLength++;
		}
	}
	
	arrFbox.sort();
	arrTbox.sort();
	fbox.length = 0;
	tbox.length = 0;
	var c;
	
	for(c = 0; c < arrFbox.length; c++){
		var no = new Option();
		no.value = arrLookup[arrFbox[c]];
		no.text = arrFbox[c];
		fbox[c] = no;
	}
	
	for(c = 0; c < arrTbox.length; c++){
		var no = new Option();
		no.value = arrLookup[arrTbox[c]];
		no.text = arrTbox[c];
		tbox[c] = no;
	}
}


function controllaCampiInput(){	
	//valorizzo il campo nascosto "ListTarget" con la lista dei Target della news separati da "|"
	var strTargets = "";
	strTargets+=listTargetCat
	if(strTargets.charAt(strTargets.length -1) == "|"){
		strTargets = strTargets.substring(0, strTargets.length -1);
	}
	
	document.form_inserisci.ListTarget.value = strTargets;
	//alert(document.form_inserisci.ListTarget.value);
	
	
	if(document.form_inserisci.num_menu.value == ""){
		alert("<%=langEditor.getTranslated("backend.categorie.detail.js.alert.insert_num_menu")%>");		
		document.form_inserisci.num_menu.focus();
		return false;
	}
	
	if(document.form_inserisci.gerarchia.value == ""){
		alert("<%=langEditor.getTranslated("backend.categorie.detail.js.alert.insert_gerarchia")%>");		
		document.form_inserisci.gerarchia.focus();
		return false;
	}
	
	if(!checkGerarchiaFormat(document.form_inserisci.gerarchia.value)){
		alert("<%=langEditor.getTranslated("backend.categorie.detail.js.alert.insert_correct_gerarchia")%>");		
		document.form_inserisci.gerarchia.focus();
		return false;
	}
	
	if(document.form_inserisci.descrizione.value == ""){
		alert("<%=langEditor.getTranslated("backend.categorie.detail.js.alert.insert_description")%>");
		document.form_inserisci.descrizione.focus();
		return false;
	}
	
	//genero il nuovo target in base al tipo di categoria specificata: contenuti o prodotti
	
	document.form_inserisci.new_target.value = document.form_inserisci.descrizione.value.toLowerCase().replace(/ /g,"_");	
	
	if(document.form_inserisci.old_target.value.toLowerCase().replace(/ /g,"_") != document.form_inserisci.new_target.value || document.form_inserisci.old_cat_type.value != document.form_inserisci.cat_type.value){
		if(confirm("<%=langEditor.getTranslated("backend.categorie.detail.js.alert.confirm_insert_target")%>")){
			document.form_inserisci.insert_new_target.value = 1
			if(document.form_inserisci.id_categoria.value == -1){
				document.form_inserisci.old_target.value = document.form_inserisci.new_target.value;
				if(confirm("<%=langEditor.getTranslated("backend.categorie.detail.js.alert.confirm_set_target_to_user")%>")){
					document.form_inserisci.set_target_to_users.value = 1
					
					if(confirm("<%=langEditor.getTranslated("backend.categorie.detail.js.alert.confirm_set_target_to_categoria")%>")){
						document.form_inserisci.set_target_to_categoria.value = 1
					}
				}
			}else{
				document.form_inserisci.old_target.value = document.form_inserisci.old_target.value.toLowerCase().replace(/ /g,"_");		
			}
		}
	}
		
	return true;
}

function replaceChars(inString){
	var outString = inString;

	for(a = 0; a < outString.length; a++){
		if(outString.charAt(a) == '"'){
			outString=outString.substring(0,a) + "&quot;" + outString.substring(a+1, outString.length);
		}
	}
	return outString;
}

function addParentCat(gerarchiaParent){
	var gerParent = gerarchiaParent.value;
	document.form_inserisci.gerarchia.value = gerParent + ".";
}


function checkGerarchiaFormat(field){
	var fieldVal = field;	
	/*alert("fieldVal: " + fieldVal);
	
	var expr1 = /^\d+,\d+$/;
	var expr2 = /^\d+$/;
			
	var expr3 = /(^\d$)|(^\d,\d$)|(^10$)|(^10,0$)/;
	var expr4 = /(^\d{4}\/([1-9]|10|11|12)$)/;
	var expr5 = /^[0-9]$/*/
	
	var expr = /(^\d+$)|(^\d+\.\d+$)|(\.\d+$)/
	var ok = expr.test(fieldVal);
	//alert("ok: " + ok);
	
	
	/*
	var exprA0 = /^\d$/;
	var exprA1 = /^\d,\d$/;
	var exprA2 = /^10$/;
	var exprA3 = /^10,0$/;
	
	ar ok = exprA0.test(fieldVal);
	alert("ok0: " + ok);	
	ok = (ok || exprA1.test(fieldVal));		
	alert("ok1: " + ok);
	ok = (ok || exprA2.test(fieldVal));
	alert("ok2: " + ok);
	ok = (ok || exprA3.test(fieldVal));
	alert("ok3: " + ok);
	
	alert("ok: " + ok);	
	*/
	
	return ok;
}

var tempX = 0;
var tempY = 0;

jQuery(document).ready(function(){
	$(document).mousemove(function(e){
	tempX = e.pageX;
	tempY = e.pageY;
	}); 
})

function showDiv(elemID){
	var element = document.getElementById(elemID);
	var jquery_id= "#"+elemID;

	element.style.left=tempX+10;
	element.style.top=tempY+10;
	$(jquery_id).show(500);
	element.style.visibility = 'visible';		
	element.style.display = "block";
}

function hideDiv(elemID){
	var element = document.getElementById(elemID);

	element.style.visibility = 'hidden';
	element.style.display = "none";
}
</script>
</head>
<body onLoad="javascript:document.form_inserisci.num_menu.focus();">
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
			<table border="0" cellspacing="0" cellpadding="0" class="principal">
		<form action="<%=Application("baseroot") & "/editor/categorie/ProcessCategoria.asp"%>" method="post" name="form_inserisci" accept-charset="UTF-8">
		  <input type="hidden" value="<%=id_categoria%>" name="id_categoria">
		   <input type="hidden" value="0" name="insert_new_target">
		  <input type="hidden" value="0" name="set_target_to_users">
		  <input type="hidden" value="0" name="set_target_to_categoria">
		  <input type="hidden" value="" name="new_target">
		  <input type="hidden" value="<%=strDescrizione%>" name="old_target">	
		  <input type="hidden" value="<%=catType%>" name="old_cat_type">			
<!--nsys-catins1-->
<!---nsys-catins1-->
			<tr> 		  		  
			  <td align="left" valign="top">
				<span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.num_menu")%></span><br/>
				<input type="text" name="num_menu" value="<%=iNumMenu%>" maxlength="2" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);">			  
			  </td>
			  <td align="center" valign="middle">&nbsp;</td>
			  <td align="left" valign="top"><span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.parent_cat")%></span><br/>
                <select name="parent_cat" class="formFieldTXT" onchange="javascript:addParentCat(this);">
                  <option value=""></option>
                  <%On Error Resume next
				Dim listaMenu, objMenu
				Set objMenu = new MenuClass
				Set listaMenu = objMenu.getCompleteMenu()
				for each zIndex in listaMenu
					if(InStrRev(iGerarchia,".",-1,1) > 0) then iGerarchiaTmp = Left(iGerarchia, (InStrRev(iGerarchia,".",-1,1)-1)) else iGerarchiaTmp = -1 end if%>
					<option value="<%=zIndex%>" <%if (strComp(zIndex, iGerarchiaTmp, 1) = 0) then response.Write("selected")%>><%=zIndex&"&nbsp;&nbsp;("&listaMenu(zIndex).getCatDescrizione()&")"%></option>
				<%next
			       Set objMenu = nothing
		
		if(Err.number <> 0)then
			'non faccio niente
		end if%>
                </select></td>			  
			  <td align="left" valign="top">&nbsp;</td>
			  <td align="left" valign="top"><span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.gerarchia")%></span><br/>
			    <input class="formFieldTXT" type="text" name="gerarchia" value="<%=iGerarchia%>" onkeypress="javascript:return isDecimal(event);"></td>
			</tr>
			<tr> 
			  <td align="left" valign="top" colspan="5" height="20">&nbsp;</td>
			</tr>
			<tr> 		  		  
			  <td align="left" valign="top"><span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.description")%></span><br/>
			    <input type="text" name="descrizione" value="<%=strDescrizione%>" class="formFieldTXT">&nbsp;<a href="#" onMouseOver="javascript:showDiv('help_desc');" class="labelForm" onmouseout="javascript:hideDiv('help_desc');">?</a>
		  <div align="left" style="z-index:1000;position:absolute;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;" id="help_desc">
		  <%=langEditor.getTranslated("backend.categorie.detail.table.label.field_help_desc")%>
		  </div></td>
			  <td align="center" valign="middle">&nbsp;</td>
			  <td align="left" valign="top"><span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.type_cat")%></span><br/>
                <select name="cat_type" class="formFieldTXTMedium">
                  <option value="<%=Application("strContentCat")%>" <%if (catType = Application("strContentCat")) then response.Write("selected")%>><%=langEditor.getTranslated("backend.categorie.detail.table.select.option.contenuti")%></option>
<!--nsys-catins2-->
                  <option value="<%=Application("strProdCat")%>" <%if (catType = Application("strProdCat")) then response.Write("selected")%>><%=langEditor.getTranslated("backend.categorie.detail.table.select.option.prodotti")%></option>
                  <option value="<%=Application("strMixedCat")%>" <%if (catType = Application("strMixedCat")) then response.Write("selected")%>><%=langEditor.getTranslated("backend.categorie.detail.table.select.option.mixed")%></option>
<!---nsys-catins2-->
               </select>
			  </td>			  
			  <td align="left" valign="top">&nbsp;</td>
			  <td align="left" valign="top"><span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.visibile")%></span><br/>
                <select name="visibile" class="formFieldTXTShort">
                  <option value="true" <%if (bolVisible) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>
                  <option value="false" <%if not(bolVisible) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>
                </select></td>
			</tr>
			<tr> 
			  <td align="left" valign="top" colspan="5" height="20">&nbsp;</td>
			</tr>
			<tr> 
			  <td align="left" valign="top">
<!--nsys-catins3-->
			<span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.contiene_prod")%></span><br/>
			<select name="contiene_prod" class="formFieldTXTShort">
			  <option value="true" <%if (bolContProd) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>
			  <option value="false" <%if not(bolContProd) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>
			</select>
<!---nsys-catins3-->			  
			  </td>
			  <td align="center" valign="middle">&nbsp;</td>
			  <td align="left" valign="top">				
				<span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.contiene_news")%></span><br/>
				<select name="contiene_news" class="formFieldTXTShort">
					<option value="true" <%if (bolContNews) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>
					<option value="false" <%if not(bolContNews) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>
				</select>				<br/>		
			  </td>
			  <td align="left" valign="top">&nbsp;</td>
			  <td align="left" valign="top">
			<span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.url_subdomain")%></span><br/>
			<input type="text" name="sub_domain_url" value="<%=sub_domain_url%>" class="formFieldTXTLong"></td>
			</tr>
			<tr> 
			  <td align="left" valign="top" colspan="5" height="20">&nbsp;</td>
			</tr>
			<tr> 
			  <td align="left" valign="top">				
			  	<span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.page_title")%></span><br/>
				<input type="text" name="page_title" value="<%=pageTitle%>" class="formFieldTXT">
			  </td>
			  <td align="center" valign="middle">&nbsp;</td>
			  <td align="left" valign="top"><span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.meta_description")%></span><br/>
			    <input type="text" name="meta_description" value="<%=metaDescription%>" class="formFieldTXT">			  </td>
			  <td align="left" valign="top">&nbsp;</td>
			  <td align="left" valign="top"><span class="labelForm"><%=langEditor.getTranslated("backend.categorie.detail.table.label.meta_keyword")%></span><br/>
			    <input type="text" name="meta_keyword" value="<%=metaKeyword%>" class="formFieldTXT"></td>
			</tr>
			<tr> 
			  <td align="left" valign="top" colspan="5" height="20">&nbsp;</td>
			</tr>
			<tr> 
			  <td align="left" valign="top" colspan="5"> 
			<input type="hidden" value="" name="ListTarget">
			<%
			Set objT = New TargetClass
			response.write(objT.renderTargetBox("listTargetCat", "targetcatbox_sx","targetcatbox_dx",langEditor.getTranslated("backend.categorie.detail.table.label.target_x_categoria"), langEditor.getTranslated("backend.categorie.detail.table.label.target_disp"), "1,2", objCatTarget, objListaTargetPerUser, false, true, langEditor))
			Set objT = Nothing
			%><br/>
			  </td>
			</tr>
			<tr> 
			  <td align="left" valign="top" colspan="5">&nbsp;</td>
			</tr>
			  <%
			Dim listaTemplate, objTemplate
			Set objTemplate = new TemplateClass
			Set listaTemplate = objTemplate.getListaTemplates()
			
			'*** RECUPERO LA LISTA DI LANGUAGE DISPONIBILI
			Dim objLangTmp, objLangList
			Set objLangTmp = New LanguageClass
			Set objLangList = objLangTmp.getListaLanguage()
			Set objLangTmp = nothing			
			%>			
			<tr> 
			  <td align="left" valign="top" class="special"><span class="labelForm"><%=langEditor.getTranslated("backend.categorie.lista.table.header.template_id")%></span><br/>
			<select name="id_template" id="id_template" class="formFieldTXT">
			  <option value="-1"></option>
			  <%	for each xIndex in listaTemplate%>
			  <option value="<%=xIndex%>" <%if (strComp(xIndex, idTemplate, 1) = 0) then response.Write("selected")%>><%=listaTemplate(xIndex).getDescrizioneTemplate()%></option>
			  <%next%>
			</select></td>
			  <td align="center" valign="middle">&nbsp;</td>
			  <td align="left" valign="top" colspan="3" class="special">
				<div id="template_lang_cat">
				<span class="labelForm"><%=langEditor.getTranslated("backend.categorie.lista.table.header.template_id_lang")%></span>
				<%
				For Each x In objLangList
					lang_code_cat = UCase(objLangList(x).getLanguageDescrizione())
					label_lang_cat = objLangList(x).getLabelDescrizione()
					id_lang_template = objSelCategoria.findLangTemplateXCategoria(lang_code_cat, false)%>
					<div style="padding-bottom:3px;">
					<select name="id_template_<%=lang_code_cat%>" id="id_template_<%=lang_code_cat%>" class="formFieldTXT">
					<option value="-1"></option>
					<%	for each xIndex in listaTemplate%>
					<option value="<%=xIndex%>" <%if (strComp(xIndex, id_lang_template, 1) = 0) then response.Write("selected")%>><%=listaTemplate(xIndex).getDescrizioneTemplate()%></option>
					<%next%>
					</select><img width="16" height="11" border="0" style="padding-left:5px;padding-right:5px;vertical-align:middle;" alt="<%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&label_lang_cat)%>" title="<%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&label_lang_cat)%>" src="/editor/img/flag/flag-<%=lang_code_cat%>.png"><%=langEditor.getTranslated("backend.lingue.lista.table.lang_label."&label_lang_cat)%>				
					</div>
				<%Next
				%>
				</div>&nbsp;
			  </td>
			</tr>
			  <%
			  Set objLangList = nothing
			  Set listaTemplate = nothing
			  Set objTemplate = nothing%>
			<script language="JavaScript">
			$('#id_template').change(function() {
				var id_template_val_ch = $('#id_template').val();
				if(id_template_val_ch!=-1){				
					$("#template_lang_cat").show();				
				}else{
					$("#template_lang_cat").hide();				
				}
			});
	
			var id_template_val = $('#id_template').val();
			if(id_template_val!=-1){
				$("#template_lang_cat").show();
			}else{
				$("#template_lang_cat").hide();			
			}
			</script>
			</form>	
			</table>
			<br/>	    
		  <input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.categorie.detail.button.inserisci.label")%>" onclick="javascript:insertCategoria();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/categorie/ListaCategorie.asp?cssClass=LCE"%>';" />
		  <br/><br/>	
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>
<%
Set objSelCategoria = nothing
Set objCategoria = nothing%>