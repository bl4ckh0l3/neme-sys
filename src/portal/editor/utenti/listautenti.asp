<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<%
'<!--nsys-usrlist1-->
%>
<!-- #include virtual="/common/include/Objects/UserGroupClass.asp" -->
<%
'<!---nsys-usrlist1-->
%>
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function deleteUtente(theForm){
	if(confirm("<%=langEditor.getTranslated("backend.utenti.lista.js.alert.delete_user")%>?")){
		theForm.delete_utente.value = "del";
		theForm.action = "<%=Application("baseroot") & "/editor/utenti/ProcessUtente.asp"%>";
		theForm.submit();
	}

}
function deleteField(theForm){
	if(confirm("<%=langEditor.getTranslated("backend.utenti.lista.js.alert.delete_field")%>?")){
		theForm.delete_field.value = "del";
		theForm.action = "<%=Application("baseroot") & "/editor/utenti/ProcessField.asp"%>";
		theForm.submit();
	}
}



function showHideDivUserField(element){
	var elementUl = document.getElementById("usrlist");
	var elementaUl = document.getElementById("ausrlist");
	var elementUf = document.getElementById("usrfield");
	var elementaUf = document.getElementById("ausrfield");

	if(element == 'usrlist'){
		elementUf.style.visibility = 'hidden';		
		elementUf.style.display = "none";
		elementaUf.className= "";
		elementUl.style.visibility = 'visible';
		elementUl.style.display = "block";
		elementaUl.className= "active";
	}else if(element == 'usrfield'){
		elementUl.style.visibility = 'hidden';
		elementUl.style.display = "none";
		elementaUl.className= "";
		elementUf.style.visibility = 'visible';		
		elementUf.style.display = "block";
		elementaUf.className= "active";
	}
}

function changeRowListData(listCounter, objtype, field){
	if(objtype=="user"){
		var active = $("#user_active_"+listCounter).val();
		var render = "";
		
		if(active==1){
			render +=('<img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" title="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.used_user")%>" alt="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.used_user")%>" hspace="5" vspace="0" border="0">');
		}else{
			render +='<a href="';
			render+="javascript:deleteUtente(document.form_lista_"+listCounter+");";
			render+='">';
			render +=('<img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" title="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.delete_user")%>" alt="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.delete_user")%>" hspace="5" vspace="0" border="0">');
			render +=('</a>');
		}
		
		$("#cancel_"+listCounter).empty();
		$("#cancel_"+listCounter).append(render);
	
		//aggiorno data modifica
		//$("#mod_data_"+listCounter).empty();
		//$("#mod_data_"+listCounter).append('<%=FormatDateTime(now(),2)%>');
	}
}

function sortUserParam(val){
	document.user_filter.order_by.value = val;
	document.user_filter.order_by_fields.value = "";
	document.user_filter.submit();	
}

function sortFilterUserParam(val){
	document.user_filter.order_by_fields.value = val;
	document.user_filter.order_by.value = "";
	document.user_filter.submit();	
}

function filterUserParam(){
	document.user_filter.submit();	
}

function showHideDivFilter(elemID){
	if ( $('#'+elemID).is(':visible')){
		$('#'+elemID).hide();
		document.user_filter.view_filter.value = 0;
	}else{
		$('#'+elemID).show();
		document.user_filter.view_filter.value = 1;
	}
}

function showHideDivFields(elemID){
	if ( $('.'+elemID).is(':visible')){
		$('.'+elemID).hide();
		document.user_filter.view_fields.value = 0;
	}else{
		$('.'+elemID).show();
		document.user_filter.view_fields.value = 1;
	}
}

function showHideDivMailBox(elemID){
	if ( $('#'+elemID).is(':visible')){
		$('#'+elemID).hide();
	}else{
		$('#'+elemID).show();
		$('#mailbox_error').empty();
		openMailbox();
	}
}

function openMailbox(){
	$.ajax({
		async: true,
		type: "GET",
		cache: false,
		url: "<%=Application("baseroot") & "/editor/utenti/include/ajaxshowmailbox.asp"%>",
		success: function(response) {
			//alert("response: "+response);
			$("#mail_text_body").empty();
			$("#mail_text_body").append(response);
		},
		error: function() {
			//alert("errorrrrrrrrrr!");
			$("#mail_text_body").empty();
			$("#mail_text_body").append("<textarea class='formFieldTXTAREAAbstract' name='mail_body'></textarea>");
		}
	});
}
</script>
</head>
<body onLoad="showHideDivUserField('<%=showTab%>')">
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LU"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">		
		<div id="tab-user-field"><a id="ausrlist" <%if(showtab="usrlist")then response.write("class=active") end if%> href="javascript:showHideDivUserField('usrlist');"><%=langEditor.getTranslated("backend.utenti.lista.table.header.label_usr_list")%></a><a id="ausrfield" <%if(showtab="usrfield")then response.write("class=active") end if%> href="javascript:showHideDivUserField('usrfield');"><%=langEditor.getTranslated("backend.utenti.lista.table.header.label_usr_field")%></a></div>
		<div id="ajaxresp" align="center" style="background-color:#FFFF00; border:1px solid #000000; color:#000000; display:none;"></div>
		<div id="usrlist" style="visibility:visible;display:block;margin:0px;padding:0px;">
			<table class="principal" border="0" cellpadding="0" cellspacing="0" align="top">
			<tr>
<!--nsys-usrlist7-->
				<td colspan="<%if(hasUserFields)then response.write(9+objListUserField.count) else response.write("9") end if%>">
<!---nsys-usrlist7-->
				<form action="<%=Application("baseroot") & "/editor/utenti/ListaUtenti.asp"%>" method="post" name="form_search" accept-charset="UTF-8">
				<input type="hidden" value="1" name="page">
				<input type="hidden" value="LU" name="cssClass">
				<input type="submit" value="<%=langEditor.getTranslated("backend.utenti.lista.label.search")%>" class="buttonForm" hspace="4">
				<input type="text" name="search_key" value="" class="formFieldTXTLong">	
				</form>
				<input type="button" value="<%=langEditor.getTranslated("backend.utenti.lista.label.fields_view_button")%>" class="buttonForm" hspace="4" onclick="javascript:showHideDivFields('filterrowfields');">
				<input type="button" value="<%=langEditor.getTranslated("backend.utenti.lista.label.filter_button")%>" class="buttonForm" hspace="4" onclick="javascript:showHideDivFilter('filterrow');">				
				</td>				
			</tr>
			<%Set objListaRuoli =  objUtente.getListaRuoli()%>
			<tr> 
				<th colspan="2">&nbsp;</th>
				<th onclick="javascript:sortUserParam(<%if(request("order_by")="1")then response.write("2") else response.write("1") end if%>);" style="cursor:pointer;text-decoration:underline;"><%=langEditor.getTranslated("backend.utenti.lista.table.header.username")%></th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.mail")%></th>
				<th onclick="javascript:sortUserParam(<%if(request("order_by")="3")then response.write("4") else response.write("3") end if%>);" style="cursor:pointer;text-decoration:underline;padding-right:80px;"><%=langEditor.getTranslated("backend.utenti.lista.table.header.role")%></th>
				<th class="upper" onclick="javascript:sortUserParam(<%if(request("order_by")="5")then response.write("6") else response.write("5") end if%>);" style="cursor:pointer;text-decoration:underline;"><%=langEditor.getTranslated("backend.utenti.lista.table.header.user_active")%></th>
				<th class="upper" onclick="javascript:sortUserParam(<%if(request("order_by")="7")then response.write("8") else response.write("7") end if%>);" style="cursor:pointer;text-decoration:underline;"><%=langEditor.getTranslated("backend.utenti.lista.table.header.public_profile")%></th>
<!--nsys-usrlist2-->
				<th style="padding-right:80px;"><%=langEditor.getTranslated("backend.utenti.lista.table.header.group")%></th>
<!---nsys-usrlist2-->
				<th class="upper"><%=langEditor.getTranslated("backend.utenti.lista.table.header.date_insert")%></th>
				
				<%
				On Error Resume next
				userFieldcount =1
				if(hasUserFields) then
					for each k in objListUserField
						Set objField = objListUserField(k)
						labelForm = objField.getDescription()
						if not(langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription())="") then labelForm = langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription())
						%>
						<th class="filterrowfields" style="cursor:pointer;text-decoration:underline;padding-right:20px;"   onclick="javascript:sortFilterUserParam(<%=k%>);"><%=labelForm%></th>
					<%Set objField = nothing
						userFieldcount=userFieldcount+1
					next
				end if

				if(Err.number<>0) then
				'response.write(Err.description)
				end if
				%>	
			</tr>
			<tr id="filterrow">
				<form action="<%=Application("baseroot") & "/editor/utenti/ListaUtenti.asp"%>" method="post" name="user_filter">
				<input type="hidden" value="usrlist" name="showtab">
				<input type="hidden" value="<%=request("order_by")%>" name="order_by">
				<input type="hidden" value="<%=request("order_by_fields")%>" name="order_by_fields">
				<input type="hidden" value="<%=request("view_filter")%>" name="view_filter">
				<input type="hidden" value="<%=request("view_fields")%>" name="view_fields">
				<td colspan="4" style="padding-left:100px;"><input type="button" value="<%=langEditor.getTranslated("backend.utenti.lista.label.filter_button_apply")%>" class="buttonForm" style="border:1px solid #000000;padding-left:10px;padding-right:10px;" hspace="4" onclick="javascript:filterUserParam();"></td>
				<td>
				<select name="rolef" class="formfieldSelect" multiple="multiple" size="2">
				<option value=""></option>
				<%for each x in objListaRuoli.Keys
					isSelected=""
					arrFilteredField = Split(request("rolef"), ",", -1, 1)
					doSelect=false
					for cf = 0 to Ubound(arrFilteredField)
						'response.write(" - x:"&x&"; -arrFilteredField(cf):"&Trim(arrFilteredField(cf))&"; -equals:"& (Cint(x)=Cint(Trim(arrFilteredField(cf)))))
						if(Cint(x)=Cint(Trim(arrFilteredField(cf))))then	
							doSelect=true
							exit for
						end if 							
					next
					if(doSelect)then							
						isSelected="selected"
					end if%>
				<option value="<%=x%>" <%if (doSelect) then response.Write(isSelected)%>><%=objListaRuoli(x)%></option>
				<%next%>
				</SELECT>
				</td>
				<td>
				<select name="activef" class="formfieldSelect">
				<option value=""></option>
				<OPTION VALUE="0" <%if (strComp("0", request("activef"), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
				<OPTION VALUE="1" <%if (strComp("1", request("activef"), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
				</SELECT></td>
				<td>
				<select name="publicf" class="formfieldSelect">
				<option value=""></option>
				<OPTION VALUE="0" <%if (strComp("0", request("publicf"), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
				<OPTION VALUE="1" <%if (strComp("1", request("publicf"), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
				</SELECT></td>
<!--nsys-usrlist8-->
				<td>&nbsp;</td>
<!---nsys-usrlist8-->
				<td>&nbsp;</td>
				<%
				Dim userFieldcount, fieldCssClass
				userFieldcount =1
				hasFieldFilterActive = false
				Set objDictFilteredFieldActive = Server.CreateObject("Scripting.Dictionary")
				if(hasUserFields) then
					for each k in objListUserField
						Set objField = objListUserField(k)%>
						<td class="filterrowfields">
						<%
						On Error Resume next
						if(Cint(objField.getTypeField())=8 OR Cint(objField.getTypeField())=1)then
							Set objFilterfieldValue = objUserField.findFieldMatchValueUnique(objField.getID())
							if(objFilterfieldValue.count>0)then
								if(request(objUserField.getFieldPrefix()&objField.getID())<>"")then
									objDictFilteredFieldActive.add objField.getID(), request(objUserField.getFieldPrefix()&objField.getID())
									hasFieldFilterActive = true
								end if%>
								<select name="<%=objUserField.getFieldPrefix()&objField.getID()%>" class="formfieldSelect" multiple size="2">
								<option value=""></option>
								<%for each x in objFilterfieldValue.Keys
									isSelected=""
									arrFilteredField = Split(objDictFilteredFieldActive(objField.getID()), ",", -1, 1)
									doSelect=false
									for cf = 0 to Ubound(arrFilteredField)
										if(x=Trim(arrFilteredField(cf)))then	
											doSelect=true
											exit for
										end if 							
									next
									if(doSelect)then							
										isSelected="selected"
									end if
									%>
									<option value="<%=x%>" <%if (doSelect) then response.Write(isSelected)%>><%=x%></option>
								<%next%>
								</SELECT>					
							<%end if
							Set objFilterfieldValue = nothing
						end if

						if(Err.number<>0) then
						'response.write(Err.description)
						end if
						%>
						</td>
						<%Set objField = nothing
						userFieldcount=userFieldcount+1
					next
				end if
				%>
				</form>
			</tr>
			<%if(request("view_filter")<>"1")then%>
			<script>
			$("#filterrow").hide();
			</script>
			<%end if%>			
				<%
				reqUserMail = null
				order_by = 1
				rolef = null
				publicf = null
				activef = null
				bolHasObj = false
				if(request("search_key")<>"")then
					reqUserMail = request("search_key")
				end if
				if(request("rolef")<>"")then
					rolef = request("rolef")
				end if
				if(request("publicf")<>"")then
					publicf = request("publicf")
				end if
				if(request("activef")<>"")then
					activef = request("activef")
				end if
				if(request("order_by")<>"")then
					order_by = request("order_by")
				end if
				On Error Resume Next
				Set objListaUtenti = objUtente.findUtente(reqUserMail, rolef, activef, publicf, 0, order_by)
				if(objListaUtenti.Count > 0) then		
					bolHasObj = true
				end if
				if Err.number <> 0 then
					bolHasObj = false
				end if				
				
				Dim intCount, tmpObjUsr
				intCount = 0
'<!--nsys-usrlist3-->				
				Dim objGroup
				Set objGroup = New UserGroupClass
'<!---nsys-usrlist3-->
								
				if(bolHasObj)then
					Set objDictMailUsr = Server.CreateObject("Scripting.Dictionary")
					Set objDictlUsrFieldsVal = Server.CreateObject("Scripting.Dictionary")
					if(hasFieldFilterActive)then
						for each k in objListaUtenti
							doRemove=true
							for each i in objDictFilteredFieldActive
								valuetmp = objUserField.findFieldMatchValue(i, k)
								arrFilteredField = Split(objDictFilteredFieldActive(i), ",", -1, 1)
								for cf = 0 to Ubound(arrFilteredField)
									'response.write(" - k:"&k&"; -valuetmp:"&valuetmp&"; -arrFilteredField(cf):"&Trim(arrFilteredField(cf)&"; -equals:"& (valuetmp=Trim(arrFilteredField(cf)))))
									if(valuetmp=Trim(arrFilteredField(cf)))then	
										doRemove=false										
										objDictlUsrFieldsVal.add i&"-"&k, valuetmp										
										exit for
									end if							
								next
							next
							if(doRemove)then							
								objListaUtenti.remove(k)
							end if
						next
					end if
					
					for each f in objListaUtenti
						objDictMailUsr.add objListaUtenti(f).getEmail(),""
					next
					
					'***************** SE È STATO IMPOSTATO UN ORDINAMENTO SUI FILTRI RIORDINO LA LISTA UTENTI IN BASE AL FILTRO SELEZIONATO
					
								
					if(request("order_by_fields")<>"")then
						order_by_fields = request("order_by_fields")
						
						Set objListaUtentiNew = Server.CreateObject("Scripting.Dictionary")					
						Set objDictOrderFields = Server.CreateObject("Scripting.Dictionary")
						
						for each k in objListaUtenti
							if(objDictlUsrFieldsVal.Exists(order_by_fields&"-"&k))then
								valuetmp = objDictlUsrFieldsVal.item(order_by_fields&"-"&k)
							else
								valuetmp = objUserField.findFieldMatchValue(order_by_fields, k)								
							end if
							objDictOrderFields.add k, valuetmp
						next
						
						Set objDictOrderFields = objUserField.SortDictionary(objDictOrderFields,2)				
						for each j in objDictOrderFields
							objListaUtentiNew.add j, objListaUtenti(CLng(j))
						next
						Set objListaUtenti = objListaUtentiNew
					end if
					
					
					
					
					Dim newsCounter, iIndex, objTmpUtenti, objTmpUtentiKey, FromUtenti, ToUtenti, Diff
					iIndex = objListaUtenti.Count
					FromUtenti = ((numPageList * itemsXpageList) - itemsXpageList)
					Diff = (iIndex - ((numPageList * itemsXpageList)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToUtenti = iIndex - Diff
					
					totPages = iIndex\itemsXpageList
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpageList <> 0) AND not ((totPages * itemsXpageList) >= iIndex)) then
						totPages = totPages +1	
					end if		
					
					Dim styleRow, styleRow2
					styleRow2 = "table-list-on"
											
					objTmpUtenti = objListaUtenti.Items
					objTmpUtentiKey=objListaUtenti.Keys		
					for newsCounter = FromUtenti to ToUtenti
						styleRow = "table-list-off"
						if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>
						<form action="<%=Application("baseroot") & "/editor/utenti/InserisciUtente.asp"%>" method="post" name="form_lista_<%=intCount%>">
						<input type="hidden" value="<%=objTmpUtentiKey(newsCounter)%>" name="id_utente">
						<input type="hidden" value="" name="delete_utente"> 
						<input type="hidden" value="LU" name="cssClass">	
						</form>		
						<tr class="<%=styleRow%>">			
						  <%Set tmpObjUsr = objTmpUtenti(newsCounter)%>
						  <td align="center" width="25"><!--nsys-demoeditusr1--><a href="javascript:document.form_lista_<%=intCount%>.submit();"><!---nsys-demoeditusr1--><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.modify_user")%>" hspace="2" vspace="0" border="0"></a></td>
							<%if(tmpObjUsr.getUserActive() = 1) then%>
							<td align="center" width="25" id="cancel_<%=intCount%>"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.used_user")%>" hspace="5" vspace="0" border="0"></td>					
							<%else%>
							<td align="center" width="25" id="cancel_<%=intCount%>"><a href="javascript:deleteUtente(document.form_lista_<%=intCount%>);"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.delete_user")%>" hspace="5" vspace="0" border="0"></a></td>										
							<%end if%>						
							<td><%=tmpObjUsr.getUserName()%></td>
							<td><%=tmpObjUsr.getEmail()%></td>	
							<td>
							<div class="ajax" id="view_ruolo_utente_<%=intCount%>" onmouseover="javascript:showHide('view_ruolo_utente_<%=intCount%>','edit_ruolo_utente_<%=intCount%>','ruolo_utente_<%=intCount%>',500, true);">
							<%if(objListaRuoli.Exists(tmpObjUsr.getRuolo())) then
							response.Write(objListaRuoli.item(tmpObjUsr.getRuolo()))
							end if%>
							</div>
							<div class="ajax" id="edit_ruolo_utente_<%=intCount%>">
							<select name="ruolo_utente" class="formfieldAjaxSelect" id="ruolo_utente_<%=intCount%>" onblur="javascript:updateField('edit_ruolo_utente_<%=intCount%>','view_ruolo_utente_<%=intCount%>','ruolo_utente_<%=intCount%>','user',<%=tmpObjUsr.getUserID()%>,2,<%=intCount%>);">		  
							<%for each x in objListaRuoli.Keys%>
							<option value="<%=x%>" <%if (tmpObjUsr.getRuolo() = x) then response.Write("selected")%>><%=objListaRuoli(x)%></option>
							<%next%>
							</SELECT>	
							</div>
							<script>
							$("#edit_ruolo_utente_<%=intCount%>").hide();
							</script>
							</td>
							<td>
							<div class="ajax" id="view_user_active_<%=intCount%>" onmouseover="javascript:showHide('view_user_active_<%=intCount%>','edit_user_active_<%=intCount%>','user_active_<%=intCount%>',500, true);">
							<%
							if (strComp("1", tmpObjUsr.getUserActive(), 1) = 0) then 
								response.Write(langEditor.getTranslated("backend.commons.yes"))
							else 
								response.Write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</div>
							<div class="ajax" id="edit_user_active_<%=intCount%>">
							<select name="user_active" class="formfieldAjaxSelect" id="user_active_<%=intCount%>" onblur="javascript:updateField('edit_user_active_<%=intCount%>','view_user_active_<%=intCount%>','user_active_<%=intCount%>','user',<%=tmpObjUsr.getUserID()%>,2,<%=intCount%>);">
							<OPTION VALUE="0" <%if (strComp("0", tmpObjUsr.getUserActive(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
							<OPTION VALUE="1" <%if (strComp("1", tmpObjUsr.getUserActive(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
							</SELECT>	
							</div>
							<script>
							$("#edit_user_active_<%=intCount%>").hide();
							</script>						
							</td>
							<td>
							<div class="ajax" id="view_public_profile_<%=intCount%>" onmouseover="javascript:showHide('view_public_profile_<%=intCount%>','edit_public_profile_<%=intCount%>','public_profile_<%=intCount%>',500, true);">
							<%
							if (strComp("1", tmpObjUsr.getPublic(), 1) = 0) then 
								response.Write(langEditor.getTranslated("backend.commons.yes"))
							else 
								response.Write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>	
							</div>
							<div class="ajax" id="edit_public_profile_<%=intCount%>">
							<select name="public_profile" class="formfieldAjaxSelect" id="public_profile_<%=intCount%>" onblur="javascript:updateField('edit_public_profile_<%=intCount%>','view_public_profile_<%=intCount%>','public_profile_<%=intCount%>','user',<%=tmpObjUsr.getUserID()%>,2,<%=intCount%>);">
							<OPTION VALUE="0" <%if (strComp("0", tmpObjUsr.getPublic(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
							<OPTION VALUE="1" <%if (strComp("1", tmpObjUsr.getPublic(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
							</SELECT>	
							</div>
							<script>
							$("#edit_public_profile_<%=intCount%>").hide();
							</script>
							</td>
	<!--nsys-usrlist4-->
							<td style="width:300px;">
							<div class="ajax" id="view_user_group_<%=intCount%>" onmouseover="javascript:showHide('view_user_group_<%=intCount%>','edit_user_group_<%=intCount%>','user_group_<%=intCount%>',500, true);">
							<%
							if not(tmpObjUsr.getGroup()="") then
							response.write(objGroup.findUserGroupByID(tmpObjUsr.getGroup()).getShortDesc())
							end if
							%>
							</div>
							<div class="ajax" id="edit_user_group_<%=intCount%>">
							<select name="user_group" class="formfieldAjaxSelect" id="user_group_<%=intCount%>" onblur="javascript:updateField('edit_user_group_<%=intCount%>','view_user_group_<%=intCount%>','user_group_<%=intCount%>','user',<%=tmpObjUsr.getUserID()%>,2,<%=intCount%>);">
							<%
							Dim objDispGroup
							On Error Resume Next
							Set objDispGroup = objGroup.getListaUserGroup()
							if(Err.number <> 0) then
							end if
							if (Instr(1, typename(objDispGroup), "dictionary", 1) > 0) then
							for each x in objDispGroup%>
							<option value="<%=x%>" <%if (tmpObjUsr.getGroup() = x) then response.Write("selected")%>><%=objDispGroup(x).getShortDesc()%></option>
							<%next
							end if%>
							</SELECT>	
							</div>
							<script>
							$("#edit_user_group_<%=intCount%>").hide();
							</script>
							</td>
	<!---nsys-usrlist4-->
							<td><%=FormatDateTime(tmpObjUsr.getInsertDate(),2)%></td>

							<%
							On Error Resume next
							userFieldcount =1
							if(hasUserFields) then
								for each k in objListUserField
									Set objField = objListUserField(k)%>
									<td class="filterrowfields">
									<%
										fieldMatchValue=""
										if not(tmpObjUsr.getUserID()="") then
											on error resume next
											Set fieldMatchValue = objUserField.findFieldMatch(objField.getID(),tmpObjUsr.getUserID())
											if (Instr(1, typename(fieldMatchValue), "dictionary", 1) > 0) then
												fieldMatchValue = fieldMatchValue.Item("value")
											end if
											if Err.number <> 0 then
												'response.write(Err.description)
												fieldMatchValue=""
											end if			
										end if%>
										<%=fieldMatchValue%>
									</td>
								<%Set objField = nothing
								userFieldcount=userFieldcount+1
								next
							end if

							if(Err.number<>0) then
							'response.write(Err.description)
							end if
							%>						

						  </tr>			
						<%intCount = intCount +1
					next
				end if
				Set objGroup = nothing
				Set tmpObjUsr = nothing
				Set objListaUtenti = nothing
				Set objUtente = Nothing
				%>
              <tr> 
		<form action="<%=Application("baseroot") & "/editor/utenti/ListaUtenti.asp"%>" method="post" name="item_x_page">
		<input type="hidden" value="usrlist" name="showtab">
<!--nsys-usrlist5-->
			<th colspan="<%if(hasUserFields)then response.write(9+objListUserField.count) else response.write("9") end if%>">
<!---nsys-usrlist5-->
				<input type="text" name="itemsList" class="formFieldTXTNumXPage" value="<%=itemsXpageList%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
				<%						
				'**************** richiamo paginazione
				urlparamfieldfilter="&itemsList="&itemsXpageList&"&showtab=usrlist&rolef="&request("rolef")&"&publicf="&request("publicf")&"&activef="&request("activef")&"&order_by="&request("order_by")&"&view_filter="&request("view_filter")&"&view_fields="&request("view_fields")&"&order_by_fields="&request("order_by_fields")
				if(hasFieldFilterActive)then
					for each i in objDictFilteredFieldActive
					urlparamfieldfilter=urlparamfieldfilter&"&"&objUserField.getFieldPrefix()&i&"="&objDictFilteredFieldActive(i)
					next
				end if				
				call PaginazioneFrontend(totPages, numPageList, strGerarchia, "/editor/utenti/ListaUtenti.asp", urlparamfieldfilter)
				%>
			</th>
		</form>
              </tr>
		<script>
		<%if(request("view_fields")<>"1")then%>
		$('.filterrowfields').hide();
		<%end if%>
		</script>
            </table>
		<%
		Set objListUserField = nothing
		Set objUserField = nothing%>	    
		<br/>
		<form action="<%=Application("baseroot") & "/editor/utenti/InserisciUtente.asp"%>" method="post" name="form_crea">
		<input type="hidden" value="-1" name="id_utente">
		<input type="hidden" value="LU" name="cssClass">
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.lista.button.inserisci.label")%>" onclick="javascript:document.form_crea.submit();" />
		&nbsp;&nbsp;
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.lista.download_excel")%>" onclick="javascript:openWinExcel('<%=Application("baseroot")&"/editor/report/create-user-excel.asp?"&urlparamfieldfilter%>','crea_excel',400,400,100,100);" />
		&nbsp;&nbsp;
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.lista.download_csv")%>" onclick="javascript:openWinExcel('<%=Application("baseroot")&"/editor/report/create-user-csv.asp?"&urlparamfieldfilter%>','crea_excel',400,400,100,100);" />		&nbsp;&nbsp;
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.lista.button.inserisci.show_mail")%>" onclick="javascript:showHideDivMailBox('mail_communication_container');" />
		</form>
		<span class="error" id="mailbox_error"><%=msgMailSend%></span>
		<div id="mail_communication_container" style="display:none;">
		<form action="<%=Application("baseroot") & "/editor/utenti/ListaUtenti.asp"%>" method="post" name="mail_communication">
		<input type="hidden" value="usrlist" name="showtab">
		<%
		bcclist=""
		for each q in objDictMailUsr
		bcclist=bcclist&q&";"
		next
		if(Trim(bcclist)<>"")then
		bcclist = Left(bcclist,Len(bcclist)-1)
		end if
		%>
		<input type="hidden" value="<%=bcclist%>" name="bcc_list">
		<input type="hidden" value="1" name="do_send_mail">		
		<b><%=langEditor.getTranslated("backend.utenti.lista.button.inserisci.subject_mail")%></b><br/><input type="text" value="" name="mail_subject" class="formFieldTXTLong"><br/><br/>
		<b><%=langEditor.getTranslated("backend.utenti.lista.button.inserisci.text_mail")%></b><br/>
		<div id="mail_text_body"></div>
		<br/><input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.lista.button.inserisci.send_mail")%>" onclick="javascript:document.mail_communication.submit();" />		
		</form>		
		</div>	
		</div>
		<div id="usrfield" style="visibility:hidden;">
			<table class="principal" border="0" cellpadding="0" cellspacing="0" align="top">
			<tr> 
				<th colspan="2">&nbsp;</th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.description")%></th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.group")%></th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.order")%></th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.type")%></th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.type_content")%></th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.required")%></th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.enabled")%></th>
				<th><%=langEditor.getTranslated("backend.utenti.lista.table.header.use_for")%></th>
			</tr>
				<%
				Dim bolHasObj
				bolHasObj = false
				intCount = 0
				iIndex = 0				

				On Error Resume Next
				Set objListaField = objUsrField.getListUserField(null,null)
				if(objListaField.Count > 0) then		
					bolHasObj = true
				end if

				if Err.number <> 0 then
					bolHasObj = false
				end if			
				
				if(bolHasObj) then
					Dim tmpObjField				
					Dim objTmpField, objTmpFieldKey, FromField, ToField
					iIndex = objListaField.Count
					FromField = ((numPageField * itemsXpageField) - itemsXpageField)
					Diff = (iIndex - ((numPageField * itemsXpageField)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToField = iIndex - Diff
					
					totPages = iIndex\itemsXpageField
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpageField <> 0) AND not ((totPages * itemsXpageField) >= iIndex)) then
						totPages = totPages +1	
					end if		
					
					styleRow2 = "table-list-on"
					
					objTmpField = objListaField.Items
					objTmpFieldKey=objListaField.Keys		
					for newsCounter = FromField to ToField
						styleRow = "table-list-off"
						if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>
						<form action="<%=Application("baseroot") & "/editor/utenti/InserisciField.asp"%>" method="post" name="form_lista_field_<%=intCount%>">
						<input type="hidden" value="<%=objTmpFieldKey(newsCounter)%>" name="id_field">
						<input type="hidden" value="" name="delete_field"> 
						<input type="hidden" value="LU" name="cssClass">	
						</form>		
						<tr class="<%=styleRow%>">				
							<%Set tmpObjField = objTmpField(newsCounter)%>
							<td align="center" width="25"><a href="javascript:document.form_lista_field_<%=intCount%>.submit();"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.modify_field")%>" hspace="2" vspace="0" border="0"></a></td>
							<td align="center" width="25"><a href="javascript:deleteField(document.form_lista_field_<%=intCount%>);"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.utenti.lista.table.alt.delete_field")%>" hspace="5" vspace="0" border="0"></a></td>										
							<td width="15%">						
							<div class="ajax" id="view_description_<%=intCount%>" onmouseover="javascript:showHide('view_description_<%=intCount%>','edit_description_<%=intCount%>','description_<%=intCount%>',500, false);"><%=tmpObjField.getDescription()%></div>
							<div class="ajax" id="edit_description_<%=intCount%>"><input type="text" class="formfieldAjax" id="description_<%=intCount%>" name="description" onmouseout="javascript:restoreField('edit_description_<%=intCount%>','view_description_<%=intCount%>','description_<%=intCount%>','user_field',<%=tmpObjField.getID()%>,1,<%=intCount%>);" value="<%=tmpObjField.getDescription()%>"></div>
							<script>
							$("#edit_description_<%=intCount%>").hide();
							</script>
							</td>
							<td width="18%">						
							<div class="ajax" id="view_id_group_<%=intCount%>" onmouseover="javascript:showHide('view_id_group_<%=intCount%>','edit_id_group_<%=intCount%>','id_group_<%=intCount%>',500, true);"><%=tmpObjField.getObjGroup().getDescription()%></div>
							<div class="ajax" id="edit_id_group_<%=intCount%>">
							<select name="id_group" class="formfieldAjaxSelect" id="id_group_<%=intCount%>" onblur="javascript:updateField('edit_id_group_<%=intCount%>','view_id_group_<%=intCount%>','id_group_<%=intCount%>','user_field',<%=tmpObjField.getID()%>,2,<%=intCount%>);">
							<%
							On Error resume next
							Set objFieldGroup = New UserFieldGroupClass
							Dim objDispFGroup
							Set objDispFGroup = objFieldGroup.getListUserFieldGroup()
							Set objFieldGroup = nothing

							if (Instr(1, typename(objDispFGroup), "dictionary", 1) > 0) then
							for each x in objDispFGroup%>
							<option value="<%=x%>" <%if (tmpObjField.getIdGroup() = x) then response.Write("selected")%>><%if not(langEditor.getTranslated("backend.utenti.detail.table.label.group."&objDispFGroup(x).getDescription()) = "") then response.write(langEditor.getTranslated("backend.utenti.detail.table.label.group."&objDispFGroup(x).getDescription())) else response.write(objDispFGroup(x).getDescription()) end if%></option>
							<%next
							end if
							Set objDispFGroup = nothing
							if(Err.number <>0)then
							'response.write(Err.description)
							end if%>
							</select>
							</div>
							<script>
							$("#edit_id_group_<%=intCount%>").hide();
							</script>
							</td>
							<td>				
							<div class="ajax" id="view_order_<%=intCount%>" onmouseover="javascript:showHide('view_order_<%=intCount%>','edit_order_<%=intCount%>','order_<%=intCount%>',500, false);"><%=tmpObjField.getOrder()%></div>
							<div class="ajax" id="edit_order_<%=intCount%>"><input type="text" class="formfieldAjaxShort" id="order_<%=intCount%>" name="order" onmouseout="javascript:restoreField('edit_order_<%=intCount%>','view_order_<%=intCount%>','order_<%=intCount%>','user_field',<%=tmpObjField.getID()%>,1,<%=intCount%>);" value="<%=tmpObjField.getOrder()%>" maxlength="3" onkeypress="javascript:return isInteger(event);"></div>
							<script>
							$("#edit_order_<%=intCount%>").hide();
							</script>
							</td>
							<td><%=objUsrField.findTypeFieldById(tmpObjField.getTypeField())%></td>
							<td><%=objUsrField.findTypeContentById(tmpObjField.getTypeContent())%></td>
							<td>
							<div class="ajax" id="view_required_<%=intCount%>" onmouseover="javascript:showHide('view_required_<%=intCount%>','edit_required_<%=intCount%>','required_<%=intCount%>',500, true);">
							<%
							if (strComp("1", tmpObjField.getRequired(), 1) = 0) then 
								response.Write(langEditor.getTranslated("backend.commons.yes"))
							else 
								response.Write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</div>
							<div class="ajax" id="edit_required_<%=intCount%>">
							<select name="required" class="formfieldAjaxSelect" id="required_<%=intCount%>" onblur="javascript:updateField('edit_required_<%=intCount%>','view_required_<%=intCount%>','required_<%=intCount%>','user_field',<%=tmpObjField.getID()%>,2,<%=intCount%>);">
							<OPTION VALUE="0" <%if (strComp("0", tmpObjField.getRequired(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
							<OPTION VALUE="1" <%if (strComp("1", tmpObjField.getRequired(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
							</SELECT>	
							</div>
							<script>
							$("#edit_required_<%=intCount%>").hide();
							</script>
							</td>
							<td>
							<div class="ajax" id="view_enabled_<%=intCount%>" onmouseover="javascript:showHide('view_enabled_<%=intCount%>','edit_enabled_<%=intCount%>','enabled_<%=intCount%>',500, true);">
							<%
							if (strComp("1", tmpObjField.getEnabled(), 1) = 0) then 
								response.Write(langEditor.getTranslated("backend.commons.yes"))
							else 
								response.Write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</div>
							<div class="ajax" id="edit_enabled_<%=intCount%>">
							<select name="enabled" class="formfieldAjaxSelect" id="enabled_<%=intCount%>" onblur="javascript:updateField('edit_enabled_<%=intCount%>','view_enabled_<%=intCount%>','enabled_<%=intCount%>','user_field',<%=tmpObjField.getID()%>,2,<%=intCount%>);">
							<OPTION VALUE="0" <%if (strComp("0", tmpObjField.getEnabled(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
							<OPTION VALUE="1" <%if (strComp("1", tmpObjField.getEnabled(), 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
							</SELECT>	
							</div>
							<script>
							$("#edit_enabled_<%=intCount%>").hide();
							</script>
							</td>	
							<td width="12%">
							<div class="ajax" id="view_use_for_<%=intCount%>" onmouseover="javascript:showHide('view_use_for_<%=intCount%>','edit_use_for_<%=intCount%>','use_for_<%=intCount%>',500, true);">
							<%
							Select case tmpObjField.getUseFor()
								Case 1
								response.Write(langEditor.getTranslated("backend.utenti.field.use_for.registration"))
								Case 2
								response.Write(langEditor.getTranslated("backend.utenti.field.use_for.purchase"))
								Case 3
								response.Write(langEditor.getTranslated("backend.utenti.field.use_for.all"))
								Case Else
							End Select
							%>
							</div>
							<div class="ajax" id="edit_use_for_<%=intCount%>">
							<select name="use_for" class="formfieldAjaxSelect" id="use_for_<%=intCount%>" onblur="javascript:updateField('edit_use_for_<%=intCount%>','view_use_for_<%=intCount%>','use_for_<%=intCount%>','user_field',<%=tmpObjField.getID()%>,2,<%=intCount%>);">
							<option value="1"<%if ("1"=tmpObjField.getUseFor()) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.utenti.field.use_for.registration")%></option>	
<!--nsys-usrlist6-->
							<option value="2"<%if ("2"=tmpObjField.getUseFor()) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.utenti.field.use_for.purchase")%></option>	
							<option value="3"<%if ("3"=tmpObjField.getUseFor()) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.utenti.field.use_for.all")%></option>	
<!---nsys-usrlist6-->
							</SELECT>	
							</div>
							<script>
							$("#edit_use_for_<%=intCount%>").hide();
							</script>
							</td>			
						</tr>				
						<%intCount = intCount +1
					next
					Set tmpObjField = nothing
					Set objListaField = nothing
				end if%>
		      <tr> 
			<form action="<%=Application("baseroot") & "/editor/utenti/ListaUtenti.asp"%>" method="post" name="item_x_page_field">
			<input type="hidden" value="usrfield" name="showtab">
			<th colspan="10">
					<input type="text" name="itemsField" class="formFieldTXTNumXPage" value="<%=itemsXpageField%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
					<%						
					'**************** richiamo paginazione
					call PaginazioneFrontend(totPages, numPageField, strGerarchia, "/editor/utenti/ListaUtenti.asp", "&itemsField="&itemsXpageField&"&showtab=usrfield")
					%>
			</th>
			</form>
		      </tr>
		    </table>
			<br/>
			<form action="<%=Application("baseroot") & "/editor/utenti/InserisciField.asp"%>" method="post" name="form_crea_field">
			<input type="hidden" value="-1" name="id_field">
			<input type="hidden" value="LU" name="cssClass">
			<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.lista.button.inserisci_field.label")%>" onclick="javascript:document.form_crea_field.submit();" />
			</form>			
		</div>
		<%
		Set objUsrField = nothing
		%>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>