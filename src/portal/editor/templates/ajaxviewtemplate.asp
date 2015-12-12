<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include file="include/init2.asp" -->
		<table border="0" cellpadding="0" cellspacing="0" class="secondary">
		<tr>
		<th><%=langEditor.getTranslated("backend.templates.view.table.label.id_template")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.templates.view.table.label.base_template")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.templates.view.table.label.elem_x_page")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.templates.view.table.label.order_by")%></th>
		</tr>
		<tr>
		<td><%=id_template%></td>
		<td class="separator">&nbsp;</td>
		<td><%=baseTemplate%></td>
		<td class="separator">&nbsp;</td>
		<td><%=elem_x_page%></td>
		<td class="separator">&nbsp;</td>
		<td>
		<%if(1 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_title_asc"))%>
		<%if(2 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_title_desc"))%>
		<%if(3 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_abstract_asc"))%>
		<%if(4 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_abstract_desc"))%>
		<%if(5 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_abstract2_asc"))%>
		<%if(6 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_abstract2_desc"))%>
		<%if(7 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_abstract3_asc"))%>
		<%if(8 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_abstract3_desc"))%>
		<%if(9 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_testo_asc"))%>
		<%if(10 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_testo_desc"))%>
		<%if(11 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_data_pub_asc"))%>
		<%if(12 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_data_pub_desc"))%>
		<%if(13 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_data_ins_asc"))%>
		<%if(14 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_data_ins_desc"))%>
		<%if(15 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_keyword_asc"))%>
		<%if(16 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_keyword_desc"))%>
		<%if(15 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_keyword_asc"))%>
		<%if(16 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_keyword_desc"))%>
<!--nsys-tmpajx1-->
		<%if(101 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_id_prodotto_asc"))%>
		<%if(102 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_id_prodotto_desc"))%>
		<%if(103 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_nome_prod_asc"))%>
		<%if(104 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_nome_prod_desc"))%>
		<%if(105 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_prezzo_asc"))%>
		<%if(106 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_prezzo_desc"))%>
		<%if(107 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_qta_disp_asc"))%>
		<%if(108 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_qta_disp_desc"))%>
		<%if(109 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_attivo_asc"))%>
		<%if(110 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_attivo_desc"))%>
		<%if(111 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_codice_prod_asc"))%>
		<%if(112 = Cint(order_by)) then response.Write(langEditor.getTranslated("backend.templates.option.label.order_codice_prod_desc"))%>	
<!---nsys-tmpajx1-->
		</td>
		</tr>		
		<%if (Instr(1, typename(objPages), "dictionary", 1) > 0) then%>
			<tr>
			<td colspan="7">
			<table border="0" align="top" cellpadding="0" cellspacing="0" class="inner-table">
			<th><%=langEditor.getTranslated("backend.templates.view.table.label.attached_pages")%></th>
			<th>&nbsp;&nbsp;<%=langEditor.getTranslated("backend.templates.view.table.label.page_priority")%></th>
			<th colspan="3">&nbsp;</th>
			<th>&nbsp;</th>
			</tr>			
			<%Dim objFilesInTemplates, pageCounter, pathTemplatePart
			pageCounter=0
			for each z in objPages.Keys
				Set objFilesInTemplates = objPages(z)
				pathTemplatePart = Application("baseroot")&Application("dir_upload_templ")&dirTemplate&"/"
				
				if(InStrRev(objFilesInTemplates.getFileName(),".inc",-1,1)>0) then pathTemplatePart=pathTemplatePart&"include/" end if
				pathTemplatePart=pathTemplatePart&objFilesInTemplates.getFileName()
				%>
				<tr id="fileparttr_<%=z%>">
				<td style="width:30%;vertical-align:top;"><%=objFilesInTemplates.getFileName()%></td>
				<td style="width:70px;text-align:center;vertical-align:top;"><%=objFilesInTemplates.getPageNum()%></td>
				<td style="width:20px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="edit_<%=pageCounter%>_<%=z%>" src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.modify_template")%>" hspace="2" vspace="0" border="0"></td>
				<td style="width:20px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="save_<%=pageCounter%>_<%=z%>" src="<%=Application("baseroot")&"/editor/img/disk.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.save_template")%>" hspace="2" vspace="0" border="0"></td>
				<td style="width:20px;text-align:center;vertical-align:top;"><img style="cursor:pointer;" id="delete_<%=pageCounter%>_<%=z%>" src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.templates.lista.table.alt.delete_template")%>" hspace="2" vspace="0" border="0"></td>
				<td><div id="show_part_<%=pageCounter%>_<%=z%>"><form accept-charset="UTF-8" method="post" action=""><textarea name="text_part_<%=pageCounter%>_<%=z%>" id="text_part_<%=pageCounter%>_<%=z%>" class="formFieldTXTAREABig"></textarea></form></div></td>
				</tr>
				<script>
				$('#show_part_<%=pageCounter%>_<%=z%>').hide();
				$('#edit_<%=pageCounter%>_<%=z%>').click(function(){ajaxTemplateFilePart('<%=pathTemplatePart%>', '', 'text_part_<%=pageCounter%>_<%=z%>', 'loadfile', 'show_part_<%=pageCounter%>_<%=z%>', '');});
				$('#save_<%=pageCounter%>_<%=z%>').click(function(){ajaxTemplateFilePart('<%=pathTemplatePart%>', $('#text_part_<%=pageCounter%>_<%=z%>').val(), 'text_part_<%=pageCounter%>_<%=z%>', 'savefilepart', 'show_part_<%=pageCounter%>_<%=z%>', '');});
				$('#delete_<%=pageCounter%>_<%=z%>').click(function(){ajaxTemplateFilePart('<%=pathTemplatePart%>', '', 'text_part_<%=pageCounter%>_<%=z%>', 'deletefilepart', 'show_part_<%=pageCounter%>_<%=z%>', '<%=z%>');});
				</script>
				<%Set objFilesInTemplates = nothing
				pageCounter=pageCounter+1				
			next
			Set objPages = nothing%>
			</table>
			</td>
			</tr>			
		<%end if%>
		</table>