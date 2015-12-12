<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<%if not(isNull(strCSS)) ANd not(strCSS = "") then%>
<link rel="stylesheet" href="<%=Application("baseroot") & strCSS%>" type="text/css">
<%end if%>
<script type="text/JavaScript" src="<%=Application("baseroot") & "/common/js/jsDraw2D_Uncompressed.js"%>"></script> 
<script language="JavaScript">
field4prodRequired = new Hashtable();

<%
'*********** recupero la lista di field abilitati e imposto in una mappa quelli obbligatori
'*********** per fare i controlli prima dell'invio del form

On Error Resume Next
Set objListField = objProdField.getListProductField(1)

for each k in objListField
	Set objF = objListField(k)
	
	if(objF.getRequired()=1)then
    if not(lang.getTranslated("portal.commons.product_field.js.alert.insert_"&objF.getDescription()) = "") then
      response.write("field4prodRequired.put('productfield"&objF.getID()&"','"&lang.getTranslated("portal.commons.product_field.js.alert.insert_"&objF.getDescription())&"');")	
    else
      response.write("field4prodRequired.put('productfield"&objF.getID()&"','"&objF.getDescription()&"');")	
    end if  
	end if
		
	Set objF = nothing
next

Set objListField = nothing
if(Err.number <> 0) then
end if
%>



var formSent = false;
function addToCarrello(theFrom,counter){	
	var sel_qta = theFrom.qta_prodotto.value;
	if(sel_qta == ""){
		alert("<%=lang.getTranslated("frontend.template_prodotto.js.alert.select_qta_prod")%>");
		return;
	}else if(isNaN(sel_qta)){
		alert("<%=lang.getTranslated("frontend.template_prodotto.js.alert.only_integer_value")%>");
		theFrom.qta_prodotto.value = "";
		return;
	}else if(sel_qta.indexOf('.') != -1){
		alert("<%=lang.getTranslated("frontend.template_prodotto.js.alert.only_integer_value")%>");
		theFrom.qta_prodotto.value = "";
		return;
	}
	
	var arrKeys = field4prodRequired.keys();	
	for(var z=0; z<arrKeys.length; z++){
		tmpKey = arrKeys[z];
		tmpValue = field4prodRequired.get(tmpKey);

    tmpFieldName = eval("document."+theFrom.name+".productfield"+counter+tmpKey.substring(12));
  
		if(tmpFieldName != undefined){
			if(tmpFieldName.value.length==0){
				alert('<%=lang.getTranslated("frontend.template_prodotto.js.alert.insert_value_for_field")%> '+tmpValue);
				return;		
			}	
		}		
	}
	
	if(formSent == false){
		formSent = true;
		theFrom.submit();
	}else{
		alert("<%=lang.getTranslated("frontend.template_prodotto.js.alert.form_already_sent")%>");
	}
}

function changeCurrency(field){
	location.href = '<%=Application("baseroot") & "/common/include/Controller.asp?gerarchia="&strGerarchia&"&page="&numPage&"&order_by="&order_by_name&"&modelPageNum="&modelPageNum&"&currency="%>'+field.value;
}

function checkMaxQtaProd(maxQtaProd, field){
	if(Number(field.value) > maxQtaProd){
		alert("<%=lang.getTranslated("frontend.template_prodotto.js.alert.exceed_qta_prod")%>");
		field.value="";
	}
}
</SCRIPT>