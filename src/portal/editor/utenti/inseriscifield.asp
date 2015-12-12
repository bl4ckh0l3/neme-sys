<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->
<!-- #include file="include/init3.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">

function insertField(){
	if(document.form_inserisci.description.value == "") {
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.insert_description")%>");
		document.form_inserisci.description.focus();
		return false;		
	}
	
	/*if(document.form_inserisci.id_group.value == "") {
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.insert_group")%>");
		document.form_inserisci.id_group.focus();
		return false;		
	}*/

	document.form_inserisci.submit();
}

function deleteGroup(field){
	var id_group = field.options[field.selectedIndex].value;
	document.form_delete_group.id_del_group.value = id_group;
	if(confirm("<%=langEditor.getTranslated("backend.utenti.lista.js.alert.delete_group")%>?")){
		document.form_delete_group.submit();
	}
}

function insertGroup(){
	var group_description = document.getElementById("group_value");
	var group_order = document.getElementById("group_order");

	document.form_inserisci_group.desc_new_group.value = group_description.value;
	document.form_inserisci_group.order_new_group.value = group_order.value;
	
	if(group_description.value == "") {
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.insert_description")%>");
		document.form_inserisci.group_value.focus();
		return false;		
	}
	
	if(group_order.value == "") {
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.insert_order")%>");
		document.form_inserisci.group_order.focus();
		return false;		
	}
	
	document.form_inserisci_group.submit();
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



function sortDropDownListByText(elem) {  
	$("select#"+elem).each(function() {  
		var selectedValue = $(this).val();  
		$(this).html($("option", $(this)).sort(function(a, b) {  
		return a.text == b.text ? 0 : a.text < b.text ? -1 : 1  
		}));  
		$(this).val(selectedValue);  
	});  
} 
</script>
</head>
<body onLoad="javascript:document.form_inserisci.description.focus();">
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">
		<form action="<%=Application("baseroot") & "/editor/utenti/ProcessField.asp"%>" method="post" name="form_inserisci">
		  <input type="hidden" value="<%=id_field%>" name="id_field">
		<tr>
		<td>	
		  <span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.description")%></span><br>
		  <input type="text" name="description" value="<%=description%>" class="formFieldTXTMedium">&nbsp;<a href="#" onMouseOver="javascript:showDiv('help_desc');" class="labelForm" onmouseout="javascript:hideDiv('help_desc');">?</a>
		  <div align="left" style="z-index:1000;position:absolute;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;" id="help_desc">
		  <%=langEditor.getTranslated("backend.utenti.detail.table.label.field_help_desc")%>
		  </div>
		  <br/><br/>	
		  
		  <div align="left" style="float:left;"><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.group")%></span><br>
			<select name="id_group" class="formFieldSelectSimple">
			<option value=""></option>
			<%if (Instr(1, typename(objDispGroup), "dictionary", 1) > 0) then
			for each x in objDispGroup%>
			<option value="<%=x%>" <%if (idGroup = x) then response.Write("selected")%>><%if not(langEditor.getTranslated("backend.utenti.detail.table.label.group."&objDispGroup(x).getDescription()) = "") then response.write(langEditor.getTranslated("backend.utenti.detail.table.label.group."&objDispGroup(x).getDescription())) else response.write(objDispGroup(x).getDescription()) end if%></option>
			<%next
			end if%>
			</select>
		  </div>
		  <div align="left" style="float:left;text-align:left;padding-top:20px;padding-right:20px;">
		  <a href="javascript:deleteGroup(document.form_inserisci.id_group);"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" title="<%=langEditor.getTranslated("backend.utenti.detail.table.alt.delete_group")%>" alt="<%=langEditor.getTranslated("backend.utenti.detail.table.alt.delete_group")%>" hspace="5" vspace="0" border="0"></a>
		  </div>
		  <div align="left" style="text-align:left;display:block;">
			<span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.insert_group")%></span><br>
			<table>
			<tr><td>
			<%=langEditor.getTranslated("backend.utenti.detail.table.label.group")%>&nbsp;<input type="text" name="group_value" id="group_value"  value="" class="formFieldTXTMedium">&nbsp;<a href="#" onMouseOver="javascript:showDiv('help_desc_group');" class="labelForm" onmouseout="javascript:hideDiv('help_desc_group');">?</a>
			  <div align="left" style="z-index:1000;position:absolute;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;" id="help_desc_group">
			  <%=langEditor.getTranslated("backend.utenti.detail.table.label.group_help_desc")%>
			  </div>
			</td></tr>
			<tr><td>
			<%=langEditor.getTranslated("backend.utenti.detail.table.label.order")%>&nbsp;<input type="text" name="group_order" id="group_order" value="" class="formFieldTXTShort" maxlength="3" onkeypress="javascript:return isInteger(event);">
			&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.detail.button.inserisci.label")%>" onclick="javascript:insertGroup();" />
		  </td></tr>
		  </table>
		  </div>
		  <br><br>
		  
		  <div align="left" style="float:left;padding-right:20px;"><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.type")%></span><br>
			<select name="id_type" id="id_type" class="formFieldSelectSimple">
			<%for each x in typeList%>
			<option VALUE="<%=x%>" <%if not(typeField = "") AND (strComp(x, typeField, 1) = 0) then response.Write("selected")%>><%if not(langEditor.getTranslated("portal.commons.user_field.type.label."&typeList(x)) = "") then response.write(langEditor.getTranslated("portal.commons.user_field.type.label."&typeList(x))) else response.write typeList(x) end if%></option>
			<%next%>
			</select>
		  </div>

			<script>
			$('#id_type').change(function() {
				var type_val_ch = $('#id_type').val();
				if(type_val_ch==1 || type_val_ch==2 || type_val_ch==3 || type_val_ch==8 || type_val_ch==9){
					$("#field_values_div").hide();
					if(type_val_ch==8 || type_val_ch==9){
						$("#max_lenght_div").hide();
					}else{
						$("#max_lenght_div").show();
					}
				}else{
					$("#field_values_div").show();
					$("#max_lenght_div").hide();
				}

				if(type_val_ch!=4){
					$("select#id_type_content option[value=5]").remove();
				}else{
					$("select#id_type_content").append($("<option></option>").attr("value",5).text("<%if not(langEditor.getTranslated("portal.commons.user_field.type_content.label.country") = "") then response.write(langEditor.getTranslated("portal.commons.user_field.type_content.label.country")) else response.write("country") end if%>"));
				}

				sortDropDownListByText("id_type_content");
			});
			</script> 
 
		  <div align="left"><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.type_content")%></span><br>
			<select name="id_type_content" id="id_type_content" class="formFieldSelectSimple">
			<%for each x in typeContentList%>
			<option VALUE="<%=x%>" <%if not(typeContent = "") AND (strComp(x, typeContent, 1) = 0) then response.Write("selected")%>><%if not(langEditor.getTranslated("portal.commons.user_field.type_content.label."&typeContentList(x)) = "") then response.write(langEditor.getTranslated("portal.commons.user_field.type_content.label."&typeContentList(x))) else response.write typeContentList(x) end if%></option>
			<%next%>
			</select>
		  </div><br>
		  <div align="left" id="field_values_div">
			  <span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.values")%></span><br>
			  <input type="text" name="field_values" value="<%=values%>" class="formFieldTXTLong">&nbsp;<a href="#" onMouseOver="javascript:showDiv('help_desc_values');" class="labelForm" onmouseout="javascript:hideDiv('help_desc_values');">?</a>
			  <div align="left" style="z-index:1000;position:absolute;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;" id="help_desc_values">
			  <%=langEditor.getTranslated("backend.utenti.detail.table.label.field_help_desc_values")%>
			  </div>
			  <br/><br/>
		  </div>
		  <div align="left" id="max_lenght_div">
			  <span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.max_lenght")%></span><br>
			  <input type="text" name="max_lenght" value="<%=maxLenght%>" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);">
			  <br/><br/>
		  </div>

		<script>
		var type_val = $('#id_type').val();
		if(type_val==1 || type_val==2 || type_val==3 || type_val==8 || type_val==9){
			$("#field_values_div").hide();
			if(type_val==8 || type_val==9){
				$("#max_lenght_div").hide();
			}else{
				$("#max_lenght_div").show();
			}
		}else{
			$("#max_lenght_div").hide();
			$("#field_values_div").show();
		}

		if(type_val!=4){
			$("select#id_type_content option[value=5]").remove();
		}

		sortDropDownListByText("id_type_content");
		</script> 
			
		  <span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.order")%></span><br>
		  <input type="text" name="order" value="<%=order%>" class="formFieldTXTShort" maxlength="3" onkeypress="javascript:return isInteger(event);">
		  <br/><br/>
		  <div align="left" style="float:left;padding-right:20px;"><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.required")%></span><br>
			<select name="required" class="formFieldTXTShort">
			<option value="0"<%if ("0"=required) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>	
			<option value="1"<%if ("1"=required) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>	
			</SELECT>&nbsp;<a href="#" onMouseOver="javascript:showDiv('help_required');" class="labelForm" onmouseout="javascript:hideDiv('help_required');">?</a>
			  <div align="left" style="z-index:1000;position:absolute;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;" id="help_required">
			  <%=langEditor.getTranslated("backend.utenti.detail.table.label.field_help_required")%>
			  </div>
		  </div>	 	
		  <div align="left" style="float:left;padding-right:20px;"><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.enabled")%></span><br>
			<select name="enabled" class="formFieldTXTShort">
			<option value="0"<%if ("0"=enabled) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.no")%></option>	
			<option value="1"<%if ("1"=enabled) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></option>	
			</SELECT>
		  </div>	 	
		  <div align="left"><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.use_for")%></span><br>
			<select name="use_for" class="formFieldTXT">
			<option value="1"<%if ("1"=useFor) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.utenti.field.use_for.registration")%></option>	
<!--nsys-usrfld1-->
			<option value="2"<%if ("2"=useFor) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.utenti.field.use_for.purchase")%></option>	
			<option value="3"<%if ("3"=useFor) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.utenti.field.use_for.all")%></option>	
<!---nsys-usrfld1-->
			</SELECT>
		  </div>
		</form>
		</td></tr>
		</table>
		<br/>
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.detail.button.inserisci.label")%>" onclick="javascript:insertField();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/utenti/ListaUtenti.asp?cssClass=LU&showtab=usrfield"%>';" />
		<br/><br/>
		<form action="<%=Application("baseroot") & "/editor/utenti/ProcessField.asp"%>" method="post" name="form_inserisci_group">
			<input type="hidden" value="ins_group" name="action">
			<input type="hidden" value="" name="desc_new_group">
			<input type="hidden" value="" name="order_new_group">
			<input type="hidden" value="<%=id_field%>" name="id_field">
		</form>	

		<form action="<%=Application("baseroot") & "/editor/utenti/ProcessField.asp"%>" method="post" name="form_delete_group">
			<input type="hidden" value="del_group" name="action">
			<input type="hidden" value="" name="id_del_group">
			<input type="hidden" value="<%=id_field%>" name="id_field">
		</form>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>
<%
Set typeList = nothing
%>