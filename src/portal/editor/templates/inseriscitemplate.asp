<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include file="include/init3.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function sendForm(){
	if(controllaCampiInput()){
		document.form_inserisci.submit();
	}else{
		return;
	}
}
function controllaCampiInput(){
	if(document.form_inserisci.dir_new_template.value == ""){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.insert_directory_name")%>");
		document.form_inserisci.dir_new_template.focus();
		return false;
	}else if(document.form_inserisci.dir_new_template.value.indexOf(" ") != -1 || document.form_inserisci.dir_new_template.value.indexOf(",") != -1 || document.form_inserisci.dir_new_template.value.indexOf(";") != -1){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.dont_use_special_char")%>");
		document.form_inserisci.dir_new_template.value = "";
		document.form_inserisci.dir_new_template.focus();
		return false;		
	}
	
	var Filecounter = document.form_inserisci.numMaxFilesToUpload.value;
	
	for (var i = 1; i <= Filecounter; i++) {
		var strIndex = eval("document.form_inserisci.fileupload_filename"+i+".value");
		strIndex = strIndex.substring(strIndex.lastIndexOf("\\")+1, strIndex.length);
		if(strIndex.indexOf(" ") != -1 || strIndex.indexOf(",") != -1 || strIndex.indexOf(";") != -1){
			alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.dont_use_special_char")%>");
			eval("document.form_inserisci.fileupload_filename"+i+".focus()");
			return false;		
		}else{					
			eval("document.form_inserisci.fileupload_filename_send_"+i+".value = strIndex");
		}
	}
	
	for (var i = 1; i <= Filecounter; i++) {
		var strIndex = eval("document.form_inserisci.fileupload_filename"+i+".value");
		strIndex = strIndex.substring(strIndex.lastIndexOf("\\")+1, strIndex.length);
		var strPosition = eval("document.form_inserisci.fileupload_position_"+i+".value");
		if(strPosition == "" && strIndex != ""){
			alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.insert_template_priority")%>");
			eval("document.form_inserisci.fileupload_position_"+i+".focus()");
			return false;
		}
	}
	
	var FilecounterInclude = document.form_inserisci.numMaxIncludes.value;
	
	for (var i = 1; i <= FilecounterInclude; i++) {
		var strIndexInclude = eval("document.form_inserisci.fileupload_include"+i+".value");
		strIndexInclude = strIndexInclude.substring(strIndexInclude.lastIndexOf("\\")+1, strIndexInclude.length);
		if(strIndexInclude.indexOf(" ") != -1 || strIndexInclude.indexOf(",") != -1 || strIndexInclude.indexOf(";") != -1){
			alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.dont_use_special_char")%>");
			eval("document.form_inserisci.fileupload_include"+i+".focus()");
			return false;		
		}else if(strIndexInclude.length>0 && strIndexInclude.indexOf(".inc") == -1){
			alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.load_only_inc")%>");
			eval("document.form_inserisci.fileupload_include"+i+".focus()");
			return false;
		}else{					
			eval("document.form_inserisci.fileupload_include_send_"+i+".value = strIndexInclude");
		}
	}
	
	<%'if(fileupload_css_filename = "") then%>	
	var strCss = document.form_inserisci.fileupload_css.value;
	strCss = strCss.substring(strCss.lastIndexOf("\\")+1, strCss.length);
	if(strCss.indexOf(" ") != -1 || strCss.indexOf(",") != -1 || strCss.indexOf(";") != -1){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.dont_use_special_char")%>");
		document.form_inserisci.fileupload_css.value = "";
		document.form_inserisci.fileupload_css.focus();
		return false;		
	}else{
		if(strCss.length > 0){
			document.form_inserisci.fileupload_css_filename.value = strCss;	
		}		
	}	
	<%'end if%>
	
	if(isNaN(document.form_inserisci.elem_x_page.value)){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.isnan_value")%>");
		document.form_inserisci.elem_x_page.focus();
		return;		
	}	

	return true;
}


function changeNumMaxFiles(){
	if(document.form_inserisci.numMaxFiles.value == ""){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.insert_value")%>");
		document.form_inserisci.numMaxFiles.focus();
		return;
	}else if(isNaN(document.form_inserisci.numMaxFiles.value)){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.isnan_value")%>");
		document.form_inserisci.numMaxFiles.focus();
		return;		
	}
	
	//location.href = "<%=Application("baseroot") & "/editor/templates/InserisciTemplate.asp?id_template="&id_template&"&numMaxFiles="%>"+document.form_inserisci.numMaxFiles.value+"&numMaxImgs=<%=numMaxImgs%>&numMaxIncludes=<%=numMaxIncludes%>&numMaxJs=<%=numMaxJs%>";

		
	//****************************************  1° PROVA CREAZIONE FIELD DINAMICAMENTE  ****************************************	
	
	var counter = document.form_inserisci.numMaxFiles.value;

	$("#cell_fileupload_filename").empty();
	
	var render ="";

	render=render+'<div style="float:left;padding-right:20px; ">';
	render=render+'<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_filename")%></span><br>';

	for(var i=1;i<=counter;i++){
		render=render+'<input type="file" name="fileupload_filename'+i+'" value="" class="formFieldTXT">';
		render=render+'<input type="hidden" name="fileupload_filename_send_'+i+'" value="" class="formFieldTXT"><br>';
	}

	render=render+'</div>';
	render=render+'<div>';
	render=render+'<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_fileposition")%></span><br>';

	for(var i=1;i<=counter;i++){
		render=render+'<input type="text" name="fileupload_position_'+i+'" class="formFieldTXTShort" value="" onkeypress="javascript:return isIntegerUnsigned(event);"><br/>';
	}		  
	render=render+'</div>';

	$("#cell_fileupload_filename").append(render);
	
	$("#numMaxFilesToUpload").attr('value', counter);
	$("#numMaxFiles").attr('value', counter);


	/*
	****************************************   2° PROVA CREAZIONE FIELD DINAMICAMENTE  ****************************************
	
	var counter = document.form_inserisci.numMaxFiles.value;

	$("#cell_fileupload_filename").empty();

	$("#cell_fileupload_filename").append('<div style="float:left;padding-right:20px;" id="div_fileupload_filename">');
	$("#div_fileupload_filename").append('<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_filename")%></span><br/>');

	for(var i=1;i<=counter;i++){
		$("#div_fileupload_filename").append($('<input type="file"/>').attr('name', "fileupload_filename"+i).attr('class', "formFieldTXT"))
		.append($('<input type="hidden"/>').attr('name', "fileupload_filename_send_"+i))
		.append('<br/>');
	}

	$("#cell_fileupload_filename").append('<div id="div_fileupload_position">')
	$("#div_fileupload_position").append('<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_fileposition")%></span><br/>');
	
	for(var i=1;i<=counter;i++){
		$("#div_fileupload_position").append($('<input type="text"/>').attr('name', "fileupload_position_"+i).attr('class', "formFieldTXTShort").attr('onkeypress', "javascript:return isIntegerUnsigned(event);"))
		.append('<br/>');
	}		  
	
	$("#numMaxFilesToUpload").attr('value', counter);
	$("#numMaxFiles").attr('value', counter);
	*/
}

function changeNumMaxIncludes(){
	if(document.form_inserisci.numMaxIncludes.value == ""){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.insert_value")%>");
		document.form_inserisci.numMaxIncludes.focus();
		return;
	}else if(isNaN(document.form_inserisci.numMaxIncludes.value)){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.isnan_value")%>");
		document.form_inserisci.numMaxIncludes.focus();
		return;		
	}
	
	//location.href = "<%=Application("baseroot") & "/editor/templates/InserisciTemplate.asp?id_template="&id_template&"&numMaxIncludes="%>"+document.form_inserisci.numMaxIncludes.value+"&numMaxFiles=<%=numMaxFiles%>&numMaxImgs=<%=numMaxImgs%>&numMaxJs=<%=numMaxJs%>";


	var counter = document.form_inserisci.numMaxIncludes.value;

	$("#cell_fileupload_include").empty();
	
	var render ="";

	render=render+'<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_includes")%></span><br>';

	for(var i=1;i<=counter;i++){
		render=render+'<input type="file" name="fileupload_include'+i+'" value="" class="formFieldTXT">';
		render=render+'<input type="hidden" name="fileupload_include_send_'+i+'" value="" class="formFieldTXT"><br>';
	}

	$("#cell_fileupload_include").append(render);
	
	$("#numMaxIncludes").attr('value', counter);

}

function changeNumMaxJs(){
	if(document.form_inserisci.numMaxJs.value == ""){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.insert_value")%>");
		document.form_inserisci.numMaxJs.focus();
		return;
	}else if(isNaN(document.form_inserisci.numMaxJs.value)){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.isnan_value")%>");
		document.form_inserisci.numMaxJs.focus();
		return;		
	}
	
	//location.href = "<%=Application("baseroot") & "/editor/templates/InserisciTemplate.asp?id_template="&id_template&"&numMaxJs="%>"+document.form_inserisci.numMaxJs.value+"&numMaxFiles=<%=numMaxFiles%>&numMaxImgs=<%=numMaxImgs%>&numMaxIncludes=<%=numMaxIncludes%>";


	var counter = document.form_inserisci.numMaxJs.value;

	$("#cell_fileupload_js").empty();
	
	var render ="";

	render=render+'<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_js")%></span><br>';

	for(var i=1;i<=counter;i++){
		render=render+'<input type="file" name="fileupload_js'+i+'" value="" class="formFieldTXT"><br>';
	}
	
	$("#cell_fileupload_js").append(render);
	
	$("#numMaxJs").attr('value', counter);

}

function changeNumMaxImgs(){
	if(document.form_inserisci.numMaxImgs.value == ""){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.insert_value")%>");
		document.form_inserisci.numMaxImgs.focus();
		return;
	}else if(isNaN(document.form_inserisci.numMaxImgs.value)){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.isnan_value")%>");
		document.form_inserisci.numMaxImgs.focus();
		return;		
	}
	
	//location.href = "<%=Application("baseroot") & "/editor/templates/InserisciTemplate.asp?id_template="&id_template&"&numMaxImgs="%>"+document.form_inserisci.numMaxImgs.value+"&numMaxFiles=<%=numMaxFiles%>&numMaxIncludes=<%=numMaxIncludes%>&numMaxJs=<%=numMaxJs%>";


	var counter = document.form_inserisci.numMaxImgs.value;

	$("#cell_fileupload_img").empty();
	
	var render ="";

	render=render+'<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_imgs")%></span><br>';

	for(var i=1;i<=counter;i++){
		render=render+'<input type="file" name="fileupload_img'+i+'" value="" class="formFieldTXT"><br>';
	}
	
	$("#cell_fileupload_img").append(render);
	
	$("#numMaxImgs").attr('value', counter);

}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<form action="<%=Application("baseroot") & "/editor/templates/ProcessTemplate.asp"%>" method="post" name="form_inserisci" enctype="multipart/form-data">
		<table cellpadding="0" cellspacing="0" height="100%" border="0" class="principal">
		  <input type="hidden" value="<%=id_template%>" name="id_template">
		  <tr>
		  <td><span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.dir_new_template")%></span><br/>
		  <%if not(Cint(id_template) = -1) then%>
		  <input type="hidden" value="<%=dir_new_template%>" name="dir_new_template">
		  (<%=dir_new_template%>)
		  <%else%>
		  <input type="text" value="" name="dir_new_template"  class="formFieldTXT">		  
		  <%end if%>
		  </td>
		  <td>&nbsp;</td>
		  <td><span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.descrizione_template")%></span><br/>
		  <input type="text" value="<%=descrizione_template%>" name="descrizione_template"  class="formFieldTXT">
		  </td>
		  </tr>
		  <tr>
		  <td colspan="3">&nbsp;</td>
		  </tr>		  
		  <tr>
		  <td><span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_css")%></span><br/>
		  <input type="hidden" value="<%=fileupload_css_filename%>" name="fileupload_css_filename">
		  <input type="file" name="fileupload_css" class="formFieldTXT">&nbsp;<%if (Cint(id_template) <> -1)  AND not(fileupload_css_filename = "") then%>&nbsp;(<%=fileupload_css_filename%>)<%end if%>
		  </td>		  
		  <td>&nbsp;</td>
		  <td><span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_base")%></span><br/>
			<select name="base_template" class="formFieldTXTShort">
			<option value="0" <%if(0 = Cint(base_template)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>
			<option value="1" <%if(1 = Cint(base_template)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>
			</select>		  
		  </td>
		  </tr>
		  <tr>
		  <td colspan="3">&nbsp;</td>
		  </tr>		  
		  <tr>
		  <td><span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.elem_x_page")%></span><br/>
		  <input type="text" value="<%=elem_x_page%>" name="elem_x_page"  maxlength="3" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);">
		  </td>		  
		  <td>&nbsp;</td>
		  <td><span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.order_by")%></span><br/>
			<select name="order_by" class="formFieldSelect">
			<option value="1" <%if(1 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_title_asc")%></option>
			<option value="2" <%if(2 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_title_desc")%></option>
			<option value="3" <%if(3 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_abstract_asc")%></option>
			<option value="4" <%if(4 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_abstract_desc")%></option>
			<option value="5" <%if(5 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_abstract2_asc")%></option>
			<option value="6" <%if(6 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_abstract2_desc")%></option>
			<option value="7" <%if(7 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_abstract3_asc")%></option>
			<option value="8" <%if(8 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_abstract3_desc")%></option>
			<option value="9" <%if(9 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_testo_asc")%></option>
			<option value="10" <%if(10 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_testo_desc")%></option>
			<option value="11" <%if(11 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_data_pub_asc")%></option>
			<option value="12" <%if(12 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_data_pub_desc")%></option>
			<option value="13" <%if(13 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_data_ins_asc")%></option>
			<option value="14" <%if(14 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_data_ins_desc")%></option>
			<option value="15" <%if(15 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_keyword_asc")%></option>
			<option value="16" <%if(16 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_keyword_desc")%></option>
<!--nsys-tmpins1-->
			<option value="101" <%if(101 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_id_prodotto_asc")%></option>
			<option value="102" <%if(102 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_id_prodotto_desc")%></option>
			<option value="103" <%if(103 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_nome_prod_asc")%></option>
			<option value="104" <%if(104 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_nome_prod_desc")%></option>
			<option value="105" <%if(105 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_prezzo_asc")%></option>
			<option value="106" <%if(106 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_prezzo_desc")%></option>
			<option value="107" <%if(107 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_qta_disp_asc")%></option>
			<option value="108" <%if(108 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_qta_disp_desc")%></option>
			<option value="109" <%if(109 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_attivo_asc")%></option>
			<option value="110" <%if(110 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_attivo_desc")%></option>
			<option value="111" <%if(111 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_codice_prod_asc")%></option>
			<option value="112" <%if(112 = Cint(order_by)) then response.Write("selected")%>><%=langEditor.getTranslated("backend.templates.option.label.order_codice_prod_desc")%></option>
<!---nsys-tmpins1-->
			</select>		  
		  </td>
		  </tr>	
		<tr>
		  <td colspan="3" class="separator"><hr/></td>
		  </tr>		  
		  <tr>
		  <td id="cell_fileupload_filename">			  
			<div style="float:left;padding-right:20px; "> 
			<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_filename")%></span><br>
			  <%
			  Dim imgCounter
			  imgCounter = 1		  
			  for y = 1 To numMaxFiles%>				
				<input type="file" name="fileupload_filename<%=imgCounter%>" class="formFieldTXT">
				<input type="hidden" value="" name="fileupload_filename_send_<%=imgCounter%>"><br>
				<%imgCounter = imgCounter +1
			  next%>
			  </div>
			<div>
			<span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_fileposition")%></span><br>
			  <%
			  imgCounter = 1		  
			  for y = 1 To numMaxFiles%>
				<input type="text" name="fileupload_position_<%=imgCounter%>" class="formFieldTXTShort" value="" onkeypress="javascript:return isIntegerUnsigned(event);"><br/>
				<%imgCounter = imgCounter+1
			  next%>		  
			 </div>
		  </td>
		  <td width="100">
		  <span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.change_num_elements")%></span><br/>
		  <input type="text" value="<%=numMaxFiles%>" name="numMaxFiles" id="numMaxFiles" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);">&nbsp;<a href="javascript:changeNumMaxFiles();"><img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="2" hspace="2" border="0" align="middle" alt="<%=langEditor.getTranslated("backend.templates.detail.table.label.change_num_files")%>"></a>
		  <input type="hidden" value="<%=numMaxFiles%>" name="numMaxFilesToUpload" id="numMaxFilesToUpload"><br>
		  </td>
		  <td><br/>(<%=langEditor.getTranslated("backend.templates.detail.table.label.help_template_files")%>)</td>
		  </tr>
		  <tr>
		  <td colspan="3" class="separator"><hr/></td>
		  </tr>		  
		  
		  <tr>
		  <td id="cell_fileupload_include">
		  <span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_includes")%></span><br>
		  <%
		  imgCounter = 1		  
		  for y = 1 To numMaxIncludes%>
		  	<input type="file" name="fileupload_include<%=imgCounter%>" class="formFieldTXT">
			<input type="hidden" value="" name="fileupload_include_send_<%=imgCounter%>"><br>
		  	<%imgCounter = imgCounter +1
		  next%>
		  </td>		  
		  <td>
		  <span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.change_num_elements")%></span><br/>		  
		  <input type="text" value="<%=numMaxIncludes%>" name="numMaxIncludes" id="numMaxIncludes" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);">&nbsp;<a href="javascript:changeNumMaxIncludes();"><img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="2" hspace="2" border="0" align="middle" alt="<%=langEditor.getTranslated("backend.templates.detail.table.label.change_num_includes")%>"></a>
		  </td>
		  <td><br/>(<%=langEditor.getTranslated("backend.templates.detail.table.label.help_template_includes")%>)</td>
		  </tr>

		  <tr>
		  <td colspan="3" class="separator"><hr/></td>
		  </tr>
		  		  
		  <tr>
		  <td id="cell_fileupload_js">
		  <span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_js")%></span><br>
		  <%
		  imgCounter = 1		  
		  for y = 1 To numMaxJs%>
		  <input type="file" name="fileupload_js<%=imgCounter%>" class="formFieldTXT"><br>
		  <%	imgCounter = imgCounter +1
		  next%></td>
		  <td>
		  <span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.change_num_elements")%></span><br/>
		  <input type="text" value="<%=numMaxJs%>" name="numMaxJs" id="numMaxJs" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);">&nbsp;<a href="javascript:changeNumMaxJs();"><img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="2" hspace="2" border="0" align="middle" alt="<%=langEditor.getTranslated("backend.templates.detail.table.label.change_num_js")%>"></a>
		  </td>
		  <td><br/>(<%=langEditor.getTranslated("backend.templates.detail.table.label.help_template_js")%>)</td>
		  </tr>
		  <tr>
		  <td colspan="3" class="separator"><hr/></td>
		  </tr>
		  <tr>
		  <td id="cell_fileupload_img">
		  <span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.template_imgs")%></span><br>
		  <%
		  imgCounter = 1		  
		  for y = 1 To numMaxImgs%>
		  <input type="file" name="fileupload_img<%=imgCounter%>" class="formFieldTXT"><br>
		  <%	imgCounter = imgCounter +1
		  next%><br/></td>
		  <td>
		  <span class="labelForm"><%=langEditor.getTranslated("backend.templates.detail.table.label.change_num_elements")%></span><br/>
		  <input type="text" value="<%=numMaxImgs%>" name="numMaxImgs" id="numMaxImgs" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);">&nbsp;<a href="javascript:changeNumMaxImgs();"><img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="2" hspace="2" border="0" align="middle" alt="<%=langEditor.getTranslated("backend.templates.detail.table.label.change_num_imgs")%>"></a>
		  </td>
		  <td><br/>(<%=langEditor.getTranslated("backend.templates.detail.table.label.help_template_imgs")%>)</td>
		  </tr>
		  </table>
		  </form><br/>
		  <input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.templates.detail.button.inserisci.label")%>" onclick="javascript:sendForm();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.templates.detail.button.annulla.label")%>" onclick="javascript:reset();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/templates/ListaTemplates.asp?cssClass=LTP"%>';" />
		  <br/><br/>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>