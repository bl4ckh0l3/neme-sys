<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include file="include/init2.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">

function insertCountry(){
	
	if(document.form_inserisci.country_code.value == ""){
		alert("<%=langEditor.getTranslated("backend.country.detail.js.alert.insert_country_code_value")%>");
		document.form_inserisci.country_code.focus();
		return;
	}
	
	if(document.form_inserisci.country_description.value == ""){
		alert("<%=langEditor.getTranslated("backend.country.detail.js.alert.insert_country_description_value")%>");
		document.form_inserisci.country_description.focus();
		return;
	}
	
	document.form_inserisci.submit()
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
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table border="0" cellspacing="0" cellpadding="0" class="principal">
		<tr><td>
		<form action="<%=Application("baseroot") & "/editor/countries/ProcessCountry.asp"%>" method="post" name="form_inserisci">
		  <input type="hidden" value="<%=id_country%>" name="id_country">	
		  <div align="left" style="float:left;"><span class="labelForm"><%=langEditor.getTranslated("backend.country.detail.table.label.country_code")%></span><br>
		  <input type="text" name="country_code" value="<%=country_code%>" class="formFieldTXT">&nbsp;&nbsp;
		  </div>	
		  <div align="left"><span class="labelForm"><%=langEditor.getTranslated("backend.country.detail.table.label.country_description")%></span><br>
		  <input type="text" name="country_description" value="<%=country_description%>" class="formFieldTXT">&nbsp;<a href="#" onMouseOver="javascript:showDiv('help_country_description');" class="labelForm" onmouseout="javascript:hideDiv('help_country_description');">?</a>
			  <div align="left" style="z-index:1000;position:absolute;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;" id="help_country_description">
			  <%=langEditor.getTranslated("backend.country.detail.table.label.country_description_help_desc")%>
			  </div>
		  </div>
		  <br/><br/>	
		  <div align="left" style="float:left;"><span class="labelForm"><%=langEditor.getTranslated("backend.country.detail.table.label.state_region_code")%></span><br>
		  <input type="text" name="state_region_code" value="<%=state_region_code%>" class="formFieldTXT">&nbsp;&nbsp;
		  </div>	
		  <div align="left"><span class="labelForm"><%=langEditor.getTranslated("backend.country.detail.table.label.state_region_description")%></span><br>
		  <input type="text" name="state_region_description" value="<%=state_region_description%>" class="formFieldTXT">&nbsp;<a href="#" onMouseOver="javascript:showDiv('help_state_region_description');" class="labelForm" onmouseout="javascript:hideDiv('help_state_region_description');">?</a>
			  <div align="left" style="z-index:1000;position:absolute;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;" id="help_state_region_description">
			  <%=langEditor.getTranslated("backend.country.detail.table.label.state_region_description_help_desc")%>
			  </div>	
		  </div>
		  <br/><br/>	
		  <div align="left" style="float:left;padding-right:10px"><span class="labelForm"><%=langEditor.getTranslated("backend.country.detail.table.label.active")%></span><br>
			<select name="active" class="formFieldTXTShort">
			<OPTION VALUE="0" <%if (strComp("0", active, 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
			<OPTION VALUE="1" <%if (strComp("1", active, 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
			</SELECT>	
		  </div>	
		  <div align="left"><span class="labelForm"><%=langEditor.getTranslated("backend.country.detail.table.label.use_for")%></span><br>
			<select name="use_for" class="formFieldTXT">
			<option value="1"<%if ("1"=use_for) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.country.use_for.registration")%></option>	
			<option value="2"<%if ("2"=use_for) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.country.use_for.purchase")%></option>	
			<option value="3"<%if ("3"=use_for) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.country.use_for.all")%></option>	
			</SELECT>
		  </div>
		  
		 <br/> 
		<%
		'***** inizializzo gli elementi per la googlemap
		strID=id_country					
		strType=3
		
		if(Cint(strID)=-1)then
			Set objGUID = new GUIDClass
			strID=objGUID.CreateNumberGUIDRandomVarLenght(7)*(-1)
			Set objGUID = nothing%>
			<input type="hidden" value="<%=strID%>" name="pregeoloc_el_id">
		<%end if%>
		<!-- #include virtual="/editor/include/localization_widget.asp" -->			  
		  
		</form>
		</td></tr>
		</table><br/>	    
		  <input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.country.detail.button.inserisci.label")%>" onclick="javascript:insertCountry();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/countries/ListaCountry.asp?cssClass=LCT"%>';" />
		<br/><br/> 		
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>