<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include file="include/init2.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">

function insertCurrency(){
	
	if(document.form_inserisci.descrizione.value == ""){
		alert("<%=langEditor.getTranslated("backend.currency.detail.js.alert.insert_descrizione_value")%>");
		document.form_inserisci.descrizione.focus();
		return;
	}

	var thisValoreProd = document.form_inserisci.valore.value;
	if(thisValoreProd == ""){
		alert("<%=langEditor.getTranslated("backend.currency.detail.js.alert.insert_valore_value")%>");
		document.form_inserisci.valore.focus();
		return;
	}else if(thisValoreProd.indexOf('.') != -1){
		alert("<%=langEditor.getTranslated("backend.prodotti.detail.js.alert.use_only_comma")%>");
		document.form_inserisci.valore.focus();
		return;		
	}
	
	document.form_inserisci.submit()
}

$(function() {
	$('#dta_referer').datepicker({
		dateFormat: 'dd/mm/yy',
		changeMonth: true,
		changeYear: true
	});
});
</script>
</head>
<body onLoad="javascript:document.form_inserisci.descrizione.focus();">
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table border="0" cellspacing="0" cellpadding="0" class="principal">
		<tr><td>
		<form action="<%=Application("baseroot") & "/editor/currency/ProcessCurrency.asp"%>" method="post" name="form_inserisci">
		  <input type="hidden" value="<%=id_currency%>" name="id_currency">
		  <span class="labelForm"><%=langEditor.getTranslated("backend.currency.detail.table.label.descrizione_currency")%></span><br>
		  <input type="text" name="descrizione" value="<%=strDescrizione%>" class="formFieldTXT">
		  <br/><br/>	
		  <div align="left"><span class="labelForm"><%=langEditor.getTranslated("backend.currency.detail.table.label.valore")%></span><br>
		  <input type="text" name="valore" value="<%=iValore%>" class="formFieldTXTMedium" onkeypress="javascript:return isDouble(event);">
		  </div>	
		  <div align="left" style="float:left;"><span class="labelForm"><%=langEditor.getTranslated("backend.currency.detail.table.label.active")%></span><br>
			<select name="attivo" class="formFieldTXTShort">
			<option value="0"<%if ("0"=iActive) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>	
			<option value="1"<%if ("1"=iActive) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>	
			</SELECT>&nbsp;&nbsp;	
		  </div>	 	
		  <div align="left"><span class="labelForm"><%=langEditor.getTranslated("backend.currency.detail.table.label.default")%></span><br>
			<!--<select name="default" class="formFieldTXTShort">
			<option value="0"<%'if ("0"=iDefault) then response.Write(" selected")%>><%'=langEditor.getTranslated("backend.commons.no")%></option>	
			<option value="1"<%'if ("1"=iDefault) then response.Write(" selected")%>><%'=langEditor.getTranslated("backend.commons.yes")%></option>	
			</SELECT>-->
			<input type="hidden" name="default" value="<%=iDefault%>">
			<%if ("0"=iDefault) then response.Write(langEditor.getTranslated("backend.commons.no")) end if%>
			<%if ("1"=iDefault) then response.Write(langEditor.getTranslated("backend.commons.yes")) end if%><br>
		  </div><br>	  	 	
		  <div align="left"><span class="labelForm"><%=langEditor.getTranslated("backend.currency.detail.table.label.dta_referer")%></span><br>
			<input type="text" name="dta_referer" id="dta_referer" value="<%=dtRefer%>" class="formFieldTXTMedium">	
		  </div>
		</form><br>
		</td></tr>
		</table>	<br> 
		  <input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.currency.detail.button.inserisci.label")%>" onclick="javascript:insertCurrency();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/currency/ListaCurrency.asp?cssClass=LCY"%>';" />
		<br/><br/>	
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>