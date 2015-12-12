<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include file="include/init2.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
/*function setDescrizione(){
	var type = document.form_inserisci.target_type.value;
	var prefix;
	
	if(type == 1)
		prefix = "<%=Application("strCatPrefix")%>";
	else if(type == 2)
		prefix = "<%=Application("strProdPrefix")%>";
	else if(type == 3)
		prefix = "<%=Application("strLangPrefix")%>";
	else
		prefix = "";
		
	document.form_inserisci.descrizione.value = prefix;
}*/

function insertTarget(){
	
	if(document.form_inserisci.target_type.value == ""){
		alert("<%=langEditor.getTranslated("backend.targets.detail.js.alert.insert_target_type")%>");
		document.form_inserisci.reset();
		document.form_inserisci.target_type.focus();
		return;
	}
	
	if(document.form_inserisci.descrizione.value == ""){
		alert("<%=langEditor.getTranslated("backend.targets.detail.js.alert.insert_target_value")%>");
		document.form_inserisci.reset();
		document.form_inserisci.descrizione.focus();
		return;
	}
	
	if(!checkTargetFormat(document.form_inserisci.descrizione.value)){
		alert("<%=langEditor.getTranslated("backend.targets.detail.js.alert.wrong_format_target_value")%>");
		document.form_inserisci.reset();
		document.form_inserisci.descrizione.focus();
		return;
	}
	/*
	if(document.form_inserisci.target_type.value == 1 && document.form_inserisci.descrizione.value.indexOf("<%=Application("strCatPrefix")%>") == -1){
		alert("<%=langEditor.getTranslated("backend.targets.detail.js.alert.insert_correct_prefix_value")%>");
		document.form_inserisci.reset();
		return;
	}else if(document.form_inserisci.target_type.value == 2 && document.form_inserisci.descrizione.value.indexOf("<%=Application("strProdPrefix")%>") == -1){
		alert("<%=langEditor.getTranslated("backend.targets.detail.js.alert.insert_correct_prefix_value")%>");
		document.form_inserisci.reset();
		return;
	}else if(document.form_inserisci.target_type.value == 3 && document.form_inserisci.descrizione.value.indexOf("<%=Application("strLangPrefix")%>") == -1){
		alert("<%=langEditor.getTranslated("backend.targets.detail.js.alert.insert_correct_prefix_value")%>");
		document.form_inserisci.reset();
		return;
	}*/
	
	document.form_inserisci.submit()
}

function checkTargetFormat(field){
	var fieldVal = field;	
	
	var expr = /^(\w|-)+$/;
	var ok = expr.test(fieldVal);	
	return ok;
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
<body onLoad="javascript:document.form_inserisci.descrizione.focus();">
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">
		<form action="<%=Application("baseroot") & "/editor/targets/ProcessTarget.asp"%>" method="post" name="form_inserisci">
		  <input type="hidden" value="<%=id_target%>" name="id_target">
			<tr><td>
		  <span class="labelForm"><%=langEditor.getTranslated("backend.target.detail.table.header.target_type")%></span><br>
		  <%
		  Dim objTTmp
		  Set objTTmp = new TargetClass
		  Set typeTarget = objTTmp.getListaTargetType()%>
			<select name="target_type" class="formFieldTXT"><!-- onChange="setDescrizione();" -->
			<option value=""></option>
			<%if not (isNull(typeTarget)) then
				for each y in typeTarget.Keys
'<!--nsys-trgins1-->
					if(y <> 3) then
'<!---nsys-trgins1-->
%>
					<option value="<%=y%>"<%if (y=iType) then response.Write(" selected")%>><%=langEditor.getTranslated(typeTarget(y))%></option>	
				<%	end if
				next
			end if%>
			</SELECT>		  
		  <%Set typeTarget = nothing
		  Set objTTmp = nothing%>
		  <br/><br/>
		  <span class="labelForm"><%=langEditor.getTranslated("backend.targets.detail.table.label.target_name")%></span><br>
		  <input type="text" name="descrizione" value="<%=strDescrizione%>" class="formFieldTXT" onkeypress="javascript:return isCorrectChar(event);">
		  <br/><br/>
		  <span class="labelForm"><%=langEditor.getTranslated("backend.targets.detail.table.label.automatic")%></span><br>
			<select name="automatic" class="formFieldTXTShort">
			<option value="0"<%if ("0"=bolAutomatic) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>	
			<option value="1"<%if ("1"=bolAutomatic) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>
			</SELECT>
		  <br/>
		</td></tr>
		</form>
		</table><br/>		    
		  <input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.target.detail.button.inserisci.label")%>" onclick="javascript:insertTarget();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/targets/ListaTarget.asp?cssClass=LT"%>';" />
		<br/><br/>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>