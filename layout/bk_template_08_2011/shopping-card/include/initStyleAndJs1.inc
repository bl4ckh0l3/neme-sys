<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<%if not(isNull(strCSS)) ANd not(strCSS = "") then%>
<link rel="stylesheet" href="<%=Application("baseroot") & strCSS%>" type="text/css">
<%end if%>
<script type="text/JavaScript" src="<%=Application("baseroot") & "/common/js/jsDraw2D_Uncompressed.js"%>"></script>
<script language="JavaScript">
function sendCarrello(){
	var id_carrello = document.form_insert_carrello.id_carrello.value;
	if(id_carrello == ""){
		alert("<%=lang.getTranslated("frontend.carrello.js.alert.no_carrello_found")%>");
		return;// false;
	}
	
	document.form_insert_carrello.submit();
}

function delFromCarrello(theForm){
	if(confirm("<%=lang.getTranslated("frontend.carrello.js.alert.confirm_del_prod")%>")){
		theForm.submit();
	}
}
</SCRIPT>