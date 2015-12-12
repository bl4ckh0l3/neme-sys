<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/UTF8Filer.asp" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file='treeviewFunctions.asp'-->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function deleteTemplate(id_objref,row,refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.templates.lista.js.alert.confirm_delete_template")%>")){
/*<!--nsys-demoedittmp1-->*/
		ajaxDeleteItem(id_objref,"template",row,refreshrows);
/*<!---nsys-demoedittmp1-->*/
		
	}
}
function confirmClone(theForm){
	if(confirm("<%=langEditor.getTranslated("backend.templates.lista.js.alert.confirm_clone_template")%>")){
/*<!--nsys-demoedittmp2-->*/
		theForm.submit();
/*<!---nsys-demoedittmp2-->*/
	}
}

function ajaxTemplateFile(path, content, textarea, command, container){
/*<!--nsys-demoedittmp3-->*/
	var dataString;
	
	//alert($('#'+container).css("display"));

	// seleziono il comando da eseguire
	switch(command) {
		case "loadfile":
			$('#message').empty();
			if($('#'+container).css("display")=="none"){
				dataString = 'filepath='+ path + '&command=' + command;  
				$.ajax({  
					type: "POST",  
					url: "<%=Application("baseroot") & "/editor/templates/ajaxtemplatefile.asp"%>",  
					data: dataString,  
					success: function(response) {  
						$('#'+textarea).val(response);			
					},
					error: function() {
						$('#message').html("<h2><%=langEditor.getTranslated("backend.commons.fail_updated_file")%></h2>")
					}
				}); 
			}else{
				$('#'+textarea).val();
			}
			$('#'+container).slideToggle();
			break; //si ferma qui

		case "savefile":		
			$('#message').empty();
			if($('#'+container).css("display")!="none"){
				if(confirm("<%=langEditor.getTranslated("backend.templates.lista.js.alert.confirm_save_file")%>")){
					dataString = 'filepath='+ path + '&command=' + command + '&content=' + encodeURIComponent(content);  
					//alert (dataString);
					$.ajax({  
						type: "POST",  
						url: "<%=Application("baseroot") & "/editor/templates/ajaxtemplatefile.asp"%>",  
						data: dataString,  
						success: function(response) {
							$('#'+container).hide();	  
							$('#message').html("<h2><%=langEditor.getTranslated("backend.commons.ok_updated_file")%></h2>"); 
							$('#'+textarea).val();			
						},
						error: function() {
							$('#message').html("<h2><%=langEditor.getTranslated("backend.commons.fail_updated_file")%></h2>");
						}
					}); 
				}
			}
			break; //si ferma qui

		default:
			//istruzioni
	}	
/*<!---nsys-demoedittmp3-->*/
	return false; 	
}

function ajaxTemplateFilePart(path, content, textarea, command, container, fileid){
/*<!--nsys-demoedittmp4-->*/
	var dataString;

	// seleziono il comando da eseguire
	switch(command) {
		case "loadfile":
			if($('#'+container).css("display")=="none"){
				dataString = 'filepath='+ path + '&command=' + command;  
				$.ajax({  
					type: "POST",  
					url: "<%=Application("baseroot") & "/editor/templates/ajaxtemplatefile.asp"%>",  
					data: dataString,  
					success: function(response) {  
						$('#'+textarea).val(response);			
					}
				}); 
			}else{
				$('#'+textarea).val();
			}
			$('#'+container).slideToggle();
			break; //si ferma qui

		case "savefilepart":	
			if($('#'+container).css("display")!="none"){
				if(confirm("<%=langEditor.getTranslated("backend.templates.lista.js.alert.confirm_save_file")%>")){	
					dataString = 'filepath='+ path + '&command=' + command + '&content=' + encodeURIComponent(content);  
					//alert (dataString);
					$.ajax({  
						type: "POST",  
						url: "<%=Application("baseroot") & "/editor/templates/ajaxtemplatefile.asp"%>",  
						data: dataString,  
						success: function(response) {
							$('#'+container).hide();	  
							$('#'+textarea).val();			
						}
					}); 
				}
			}
			break; //si ferma qui

		case "deletefilepart":
			if(confirm("<%=langEditor.getTranslated("backend.templates.lista.js.alert.confirm_delete_file")%>")){			
				dataString = 'filepath='+ path + '&command=' + command + '&fileid='+fileid;  
				//alert (dataString);
				$.ajax({  
					type: "POST",  
					url: "<%=Application("baseroot") & "/editor/templates/ajaxtemplatefile.asp"%>",  
					data: dataString,  
					success: function(response) {
						$('#'+container).hide();	  
						$('#'+textarea).val();
						$('#fileparttr_'+fileid).remove();					
					}
				});
			} 
			break; //si ferma qui

		default:
			//istruzioni
	}	
/*<!---nsys-demoedittmp4-->*/
	return false; 	
}

function ajaxViewZoom(id_template, container){
	var dataString;

	if($('#'+container).css("display")=="none"){
		dataString = 'id_template='+ id_template;  
		$.ajax({  
			type: "POST",  
			url: "<%=Application("baseroot") & "/editor/templates/ajaxviewtemplate.asp"%>",  
			data: dataString,  
			success: function(response) {  
				$('#'+container).html(response); 			
			}
		}); 
	}else{
		$('#'+container).empty();
	}
	$('#'+container).slideToggle();	

	return false; 	
}

/*function flip(l){
	if (document.getElementById){
		var on = (document.getElementById(l).style.display == 'none') ? 1 : 0;
		document.getElementById(l).style.display = (on) ? 'block' : 'none';
		document.images['i'+l].src = (on) ? 'img/minus.gif' : 'img/plus.gif';
	}
}*/
</script>

 <!--<STYLE type="text/css">
.treeview * { font-size: 12px; font-family:tahoma; color:black }
.treeview a { text-decoration:none;color:black }
.treeview a:hover { text-decoration:underline;color:blue }
.treeview div { padding:2px }
.treeview img.r0 { margin-left: 10px }
.treeview img.r0a { margin-left: 36px }
.treeview img.r1 { margin-left: 20px }
.treeview img.r1a { margin-left: 46px }
.treeview img.r2 { margin-left: 30px }
.treeview img.r2a { margin-left: 56px }
 </STYLE>-->
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LTP"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">
			<tr nowrap> 
				<th colspan="4">&nbsp;</th>
				<th><%=langEditor.getTranslated("backend.templates.lista.table.header.descrizione")%></th>
				<th><%=langEditor.getTranslated("backend.templates.lista.table.header.template_dir")%></th>
				<th><%=langEditor.getTranslated("backend.templates.lista.table.header.template_css")%></th>
			</tr>			  
				<%
				Dim hasTemplate
				hasTemplate = false
				on error Resume Next
					Set objListaTemplates = objTemplates.getListaTemplates()	
					
					if(objListaTemplates.Count > 0) then
						hasTemplate = true
					end if
					
				if Err.number <> 0 then
					'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
					hasTemplate = false
				end if	
				
				if(hasTemplate) then	
									
					Dim intCount
					intCount = 0
					
					Dim templCounter, iIndex, objTmpTempl, objTmpTemplKey, FromTempl, ToTempl, Diff
					iIndex = objListaTemplates.Count
					FromTempl = ((numPage * itemsXpage) - itemsXpage)
					Diff = (iIndex - ((numPage * itemsXpage)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToTempl = iIndex - Diff
					
					totPages = iIndex\itemsXpage
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
						totPages = totPages +1	
					end if		
							
					objTmpTempl = objListaTemplates.Items
					objTmpTemplKey=objListaTemplates.Keys
					
					Dim styleRow, styleRow2
					styleRow2 = "table-list-on"
					
							
					for templCounter = FromTempl to ToTempl
						styleRow = "table-list-off"
						if(templCounter MOD 2 = 0) then styleRow = styleRow2 end if
						Set objFilteredTempl = objTmpTempl(templCounter)
						%>
						<form action="<%=Application("baseroot") & "/editor/templates/InserisciTemplate.asp"%>" method="post" name="form_lista_<%=intCount%>">
						<input type="hidden" value="<%=objFilteredTempl.getID()%>" name="id_template">
						<input type="hidden" value="<%=cssClass%>" name="cssClass">
						</form>	
						<tr class="<%=styleRow%>" id="tr_delete_list_<%=intCount%>">
						<td align="center" width="25"><img style="cursor:pointer;" id="clone_zoom_<%=intCount%>" src="<%=Application("baseroot")&"/editor/img/page_white_copy.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.clone")%>" hspace="2" vspace="0" border="0"></td>
						<td align="center" width="25"><img style="cursor:pointer;" id="view_zoom_<%=intCount%>" src="<%=Application("baseroot")&"/editor/img/zoom.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.view")%>" hspace="2" vspace="0" border="0"></td>
						<td align="center" width="25"><a href="javascript:document.form_lista_<%=intCount%>.submit();"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.modify_template")%>" hspace="2" vspace="0" border="0"></a></td>	
						<%if(objFilteredTempl.getBaseTemplate = 0) then%>
						<td align="center" width="25"><a href="javascript:deleteTemplate(<%=objFilteredTempl.getID()%>, 'tr_delete_list_<%=intCount%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.delete_template")%>" hspace="2" vspace="0" border="0"></a></td>
						<%else%>
						<td align="center" width="25"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.cant_delete")%>" hspace="2" vspace="0" border="0"></td>
						<%end if%>
						<td nowrap width="30%">					
						<div class="ajax" id="view_descrizione_template_<%=intCount%>" onmouseover="javascript:showHide('view_descrizione_template_<%=intCount%>','edit_descrizione_template_<%=intCount%>','descrizione_template_<%=intCount%>',500, false);"><%=objFilteredTempl.getDescrizioneTemplate()%></div>
						<div class="ajax" id="edit_descrizione_template_<%=intCount%>"><input type="text" class="formfieldAjax" id="descrizione_template_<%=intCount%>" name="descrizione_template" onmouseout="javascript:restoreField('edit_descrizione_template_<%=intCount%>','view_descrizione_template_<%=intCount%>','descrizione_template_<%=intCount%>','template',<%=objFilteredTempl.getID()%>,1,<%=intCount%>);" value="<%=objFilteredTempl.getDescrizioneTemplate()%>"></div>
						<script>
						$("#edit_descrizione_template_<%=intCount%>").hide();
						</script>
						</td>
						<td nowrap width="30%"><%=objFilteredTempl.getDirTemplate()%></td>
						<td><%=objFilteredTempl.getTemplateCss()%></td>
						</tr>
						
						<tr class="preview_row">
						<td colspan="8">
						<div id="view_template_<%=intCount%>"></div>						
						<div id="clone_template_<%=intCount%>">
						<form action="<%=Application("baseroot") & "/editor/templates/clonetemplate.asp"%>" method="post" name="form_clone_<%=intCount%>">
						<input type="hidden" value="<%=objFilteredTempl.getID()%>" name="id_template">
						<input type="text" class="formfieldAjax" name="new_dir_template">
						<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.templates.lista.button.label.inserisci_dir")%>" onclick="javascript:confirmClone(document.form_clone_<%=intCount%>);" />
						</form>
						</div>
						<script>
						$("#view_template_<%=intCount%>").hide();
						$("#clone_template_<%=intCount%>").hide();
						$('#view_zoom_<%=intCount%>').click(function(){$('#clone_template_<%=intCount%>').hide();ajaxViewZoom('<%=objFilteredTempl.getID()%>', 'view_template_<%=intCount%>');});
						$('#clone_zoom_<%=intCount%>').click(function(){$('#view_template_<%=intCount%>').hide();$('#clone_template_<%=intCount%>').slideToggle();});
						</script>
						</td>
						</tr>
						
						<%intCount = intCount +1
						Set objFilteredTempl = nothing
					next
					Set objListaTemplates = nothing
					%>
				  
				  <tr> 
					<form action="<%=Application("baseroot") & "/editor/templates/ListaTemplates.asp"%>" method="post" name="item_x_page">
					<th colspan="7">
					<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
					<%		
					'**************** richiamo paginazione
					call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/templates/ListaTemplates.asp", "&order_by="&order_templ_by&"&items="&itemsXpage)%>
					</th
					></form>
              	</tr>
              	<%end if
				Set objTemplates = Nothing%>
		</table>
		<br/>
		<form action="<%=Application("baseroot") & "/editor/templates/InserisciTemplate.asp"%>" method="post" name="form_crea">
		<input type="hidden" value="LTP" name="cssClass">
		<input type="hidden" value="-1" name="id_template">
		<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.templates.lista.button.label.inserisci")%>" onclick="javascript:document.form_crea.submit();" />
		&nbsp;&nbsp;<a href="<%=Application("baseroot") & "/public/utils/templates_neme-sys.zip"%>"><img src="<%=Application("baseroot")&"/common/img/iconaZip.gif"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.download_template_example")%>" hspace="2" vspace="0" border="0" align="absmiddle"></a>
		</form>
		
		<div class="treeview">	
		<%Dim treeRoot, elemId, elementsMap, rowCounter%>		
		<span id="message"></span>
		<br/>
		<table class="secondary" border="0" align="top" cellpadding="0" cellspacing="0">
			<tr nowrap>
			<th style="text-align:center;width:25px;vertical-align:top;"><img style="cursor:pointer;" id="view_widget_files" src="<%=Application("baseroot")&"/editor/img/zoom.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.view")%>" hspace="2" vspace="0" border="0"></th>
			<th style="text-align:left;vertical-align:top;"><%=langEditor.getTranslated("backend.templates.lista.table.header.widget_files")%></th>
		    </tr>
		</table>
		<div id="widget_files">
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">
			<%
			intCount = 0
			prefix="-wf-"
			Set elementsMap = CreateObject("Scripting.Dictionary")
			treeRoot = Application("baseroot") & "/public/layout/addson/"
			call generateTree(treeRoot, elementsMap)	

			for each elem in elementsMap
				elemId = Mid(elem,1,InStrRev(elem,".",-1,1)-1)&prefix&intCount %>			
				<tr nowrap> 
					<td style="width:25px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="edit_<%=elemId%>" src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.modify_template")%>" hspace="2" vspace="0" border="0"></td>
					<td style="width:25px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="save_<%=elemId%>" src="<%=Application("baseroot")&"/editor/img/disk.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.save_template")%>" hspace="2" vspace="0" border="0"></td>
					<td style="text-align:left;vertical-align:top;width:25%;" valign="top"><img src="img/folder.gif" hspace="2" border="0" />&nbsp;<%=elementsMap(elem)%></td>
					<td style="text-align:left;vertical-align:top;width:25%;" valign="top"><img src="img/file.gif" hspace="2" border="0" />&nbsp;<%=elem%></td>
					<td><div id="show_<%=elemId%>"><form accept-charset="UTF-8" method="post" action=""><textarea name="text_<%=elemId%>" id="text_<%=elemId%>" class="formFieldTXTAREABig"></textarea></form></div></td>
				</tr>
				<script>
				$('#show_<%=elemId%>').hide();
				$('#edit_<%=elemId%>').click(function(){ajaxTemplateFile('<%=elementsMap(elem)&elem%>', '', 'text_<%=elemId%>', 'loadfile', 'show_<%=elemId%>');});
				$('#save_<%=elemId%>').click(function(){ajaxTemplateFile('<%=elementsMap(elem)&elem%>', $('#text_<%=elemId%>').val(), 'text_<%=elemId%>', 'savefile', 'show_<%=elemId%>');});
				</script>
			<%intCount = intCount +1
			next%>
		</table>
		</div>
		<script>
		$('#widget_files').hide();
		$('#view_widget_files').click(function(){$('#widget_files').slideToggle();});
		</script>
		<table class="secondary" border="0" align="top" cellpadding="0" cellspacing="0">
			<tr nowrap> 
			<th style="text-align:center;width:25px;vertical-align:top;"><img style="cursor:pointer;" id="view_area_user_files" src="<%=Application("baseroot")&"/editor/img/zoom.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.view")%>" hspace="2" vspace="0" border="0"></th>
			<th style="text-align:left;vertical-align:top;"><%=langEditor.getTranslated("backend.templates.lista.table.header.area_user_files")%></th>
		    </tr>
		</table>
		<div id="area_user_files">
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">	
			<%
			intCount = 0
			prefix="-uf-"
			Set elementsMap = CreateObject("Scripting.Dictionary")
			treeRoot = Application("baseroot") & "/public/layout/area_user/"
			call generateTree(treeRoot, elementsMap)

			for each elem in elementsMap
				elemId = Mid(elem,1,InStrRev(elem,".",-1,1)-1)&prefix&intCount %>			
				<tr nowrap> 
					<td style="width:25px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="edit_<%=elemId%>" src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.modify_template")%>" hspace="2" vspace="0" border="0"></td>
					<td style="width:25px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="save_<%=elemId%>" src="<%=Application("baseroot")&"/editor/img/disk.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.save_template")%>" hspace="2" vspace="0" border="0"></td>
					<td style="text-align:left;vertical-align:top;width:25%;" valign="top"><img src="img/folder.gif" hspace="2" border="0" />&nbsp;<%=elementsMap(elem)%></td>
					<td style="text-align:left;vertical-align:top;width:25%;" valign="top"><img src="img/file.gif" hspace="2" border="0" />&nbsp;<%=elem%></td>
					<td><div id="show_<%=elemId%>"><form accept-charset="UTF-8" method="post" action=""><textarea name="text_<%=elemId%>" id="text_<%=elemId%>" class="formFieldTXTAREABig"></textarea></form></div></td>
				</tr>
				<script>
				$('#show_<%=elemId%>').hide();
				$('#edit_<%=elemId%>').click(function(){ajaxTemplateFile('<%=elementsMap(elem)&elem%>', '', 'text_<%=elemId%>', 'loadfile', 'show_<%=elemId%>');});
				$('#save_<%=elemId%>').click(function(){ajaxTemplateFile('<%=elementsMap(elem)&elem%>', $('#text_<%=elemId%>').val(), 'text_<%=elemId%>', 'savefile', 'show_<%=elemId%>');});
				</script>
			<%intCount = intCount +1
			next%>
		</table>
		</div>
		<script>
		$('#area_user_files').hide();
		$('#view_area_user_files').click(function(){$('#area_user_files').slideToggle();});
		</script>
		<table class="secondary" border="0" align="top" cellpadding="0" cellspacing="0">
			<tr nowrap> 
			<th style="text-align:center;width:25px;vertical-align:top;"><img style="cursor:pointer;" id="view_css_files" src="<%=Application("baseroot")&"/editor/img/zoom.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.view")%>" hspace="2" vspace="0" border="0"></th>
			<th style="text-align:left;vertical-align:top;"><%=langEditor.getTranslated("backend.templates.lista.table.header.css_files")%></th>
		    </tr>
		</table>
		<div id="css_files">
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">	
			<%
			intCount = 0
			prefix="-cf-"
			Set elementsMap = CreateObject("Scripting.Dictionary")
			treeRoot = Application("baseroot") & "/public/layout/css/"
			call generateTree(treeRoot, elementsMap)

			for each elem in elementsMap
				elemId = Mid(elem,1,InStrRev(elem,".",-1,1)-1)&prefix&intCount %>			
				<tr nowrap> 
					<td style="width:25px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="edit_<%=elemId%>" src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.modify_template")%>" hspace="2" vspace="0" border="0"></td>
					<td style="width:25px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="save_<%=elemId%>" src="<%=Application("baseroot")&"/editor/img/disk.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.save_template")%>" hspace="2" vspace="0" border="0"></td>
					<td style="text-align:left;vertical-align:top;width:25%;" valign="top"><img src="img/folder.gif" hspace="2" border="0" />&nbsp;<%=elementsMap(elem)%></td>
					<td style="text-align:left;vertical-align:top;width:25%;" valign="top"><img src="img/file.gif" hspace="2" border="0" />&nbsp;<%=elem%></td>
					<td><div id="show_<%=elemId%>"><form accept-charset="UTF-8" method="post" action=""><textarea name="text_<%=elemId%>" id="text_<%=elemId%>" class="formFieldTXTAREABig"></textarea></form></div></td>
				</tr>
				<script>
				$('#show_<%=elemId%>').hide();
				$('#edit_<%=elemId%>').click(function(){ajaxTemplateFile('<%=elementsMap(elem)&elem%>', '', 'text_<%=elemId%>', 'loadfile', 'show_<%=elemId%>');});
				$('#save_<%=elemId%>').click(function(){ajaxTemplateFile('<%=elementsMap(elem)&elem%>', $('#text_<%=elemId%>').val(), 'text_<%=elemId%>', 'savefile', 'show_<%=elemId%>');});
				</script>
			<%intCount = intCount +1
			next%>
		</table>
		</div>
		<script>
		$('#css_files').hide();
		$('#view_css_files').click(function(){$('#css_files').slideToggle();});
		</script>
		<table class="secondary" border="0" align="top" cellpadding="0" cellspacing="0">
			<tr nowrap> 
			<th style="text-align:center;width:25px;vertical-align:top;"><img style="cursor:pointer;" id="view_include_files" src="<%=Application("baseroot")&"/editor/img/zoom.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.view")%>" hspace="2" vspace="0" border="0"></th>
			<th style="text-align:left;vertical-align:top;"><%=langEditor.getTranslated("backend.templates.lista.table.header.include_files")%></th>
		    </tr>
		</table>
		<div id="include_files">
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">	
			<%
			intCount = 0
			prefix="-if-"
			Set elementsMap = CreateObject("Scripting.Dictionary")
			treeRoot = Application("baseroot") & "/public/layout/include/"
			call generateTree(treeRoot, elementsMap)

			for each elem in elementsMap
				elemId = Mid(elem,1,InStrRev(elem,".",-1,1)-1)&prefix&intCount %>			
				<tr nowrap> 
					<td style="width:25px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="edit_<%=elemId%>" src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.modify_template")%>" hspace="2" vspace="0" border="0"></td>
					<td style="width:25px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="save_<%=elemId%>" src="<%=Application("baseroot")&"/editor/img/disk.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.save_template")%>" hspace="2" vspace="0" border="0"></td>
					<td style="text-align:left;vertical-align:top;width:25%;" valign="top"><img src="img/folder.gif" hspace="2" border="0" />&nbsp;<%=elementsMap(elem)%></td>
					<td style="text-align:left;vertical-align:top;width:25%;" valign="top"><img src="img/file.gif" hspace="2" border="0" />&nbsp;<%=elem%></td>
					<td><div id="show_<%=elemId%>"><form accept-charset="UTF-8" method="post" action=""><textarea name="text_<%=elemId%>" id="text_<%=elemId%>" class="formFieldTXTAREABig"></textarea></form></div></td>
				</tr>
				<script>
				$('#show_<%=elemId%>').hide();
				$('#edit_<%=elemId%>').click(function(){ajaxTemplateFile('<%=elementsMap(elem)&elem%>', '', 'text_<%=elemId%>', 'loadfile', 'show_<%=elemId%>');});
				$('#save_<%=elemId%>').click(function(){ajaxTemplateFile('<%=elementsMap(elem)&elem%>', $('#text_<%=elemId%>').val(), 'text_<%=elemId%>', 'savefile', 'show_<%=elemId%>');});
				</script>
			<%intCount = intCount +1
			next%>
		</table>
		</div>
		<script>
		$('#include_files').hide();
		$('#view_include_files').click(function(){$('#include_files').slideToggle();});
		</script>
		</div>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>
