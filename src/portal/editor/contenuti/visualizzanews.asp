<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/CommentsClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include file="include/init2.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table border="0" cellpadding="0" cellspacing="0" class="principal">
		<tr>
		<th colspan="3"><%=langEditor.getTranslated("backend.news.view.table.label.title")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("portal.templates.commons.label.see_comments_news")%></th>
		</tr>
		<tr>
		<td colspan="3"><%=strTitolo%><br></td>
		<td class="separator">&nbsp;</td>
		<td>
		<%
		Set objCommento = New CommentsClass
		if(not(isNull(objCommento.findCommentiByIDElement(objSelNews.getNewsID(),1,1)))) then%>
			<a href="javascript:openWin('<%=Application("baseroot")&"/public/layout/include/popupComments.asp?id_element="&objSelNews.getNewsID()&"&element_type=1&active=1"%>','popupallegati',420,400,100,100);" title="<%=langEditor.getTranslated("portal.templates.commons.label.see_comments_news")%>"><img src="<%=Application("baseroot")&"/editor/img/comments.png"%>" hspace="0" vspace="0" border="0"></a>
		<%else
			response.Write("<div align='left'>"&langEditor.getTranslated("backend.news.detail.table.label.no_comments")&"</div>")
		end if
		Set objCommento = nothing
		%>
		</td>
		</tr>
		<tr>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.abstract_field")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.abstract_field")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.abstract_field")%></th>
		</tr>
		<tr>
		<td><%=strAbs%></td>
		<td class="separator">&nbsp;</td>
		<td><%=strAbs2%></td>
		<td class="separator">&nbsp;</td>
		<td><%=strAbs3%></td>
		</tr>
		<tr>
		<th colspan="5"><%=langEditor.getTranslated("backend.news.view.table.label.text")%></th>
		</tr>
		<tr>
		<td colspan="5"><%=strText%></td>
		</tr>
		<tr>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.keyword")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.commons.label.localization")%></th>
		<td class="separator">&nbsp;</td>
		<th></th>
		</tr>
		<tr>
		<td><%=strKeyword%></td>
		<td class="separator">&nbsp;</td>
		<td><%=strGeolocal%></td>
		<td class="separator">&nbsp;</td>
		<td>&nbsp;</td>
		</tr>
		<tr>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.page_title")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.meta_description")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.meta_keyword")%></th>
		</tr>
		<tr>
		<td><%=page_title%></td>
		<td class="separator">&nbsp;</td>
		<td><%=meta_description%></td>
		<td class="separator">&nbsp;</td>
		<td><%=meta_keyword%></td>
		</tr>		
		<tr>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.stato_news")%></th>		
		<td class="separator">&nbsp;</td>  
		<th><%=langEditor.getTranslated("backend.news.view.table.label.target_x_news")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.attached_files")%></th>
		</tr>
		<tr>
		<td><%
		Select Case stato_news
		Case 0
			response.write(langEditor.getTranslated("backend.news.view.table.label.da_editare"))
		Case 1
			response.write(langEditor.getTranslated("backend.news.view.table.label.pubblicata"))
		Case Else
		End Select
		%></td>		
		<td class="separator">&nbsp;</td>  
		<td><%		
		if (Instr(1, typename(objTarget), "dictionary", 1) > 0) then
			for each y in objTarget.Keys
				if (objTarget(y).getTargetType() = 3) then		
					if not(langEditor.getTranslated("portal.header.label.desc_lang."&Replace(objTarget(y).getTargetDescrizione(), "lang_", "", 1, -1, 1)) = "") then response.write (langEditor.getTranslated("portal.header.label.desc_lang."&Replace(objTarget(y).getTargetDescrizione(), "lang_", "", 1, -1, 1)) & "<br>") else response.write(objTarget(y).getTargetDescrizione()& "<br>") end if									
				end if		
			next		
			
			Dim CategoriatmpClass, objCategorieXCont
			Set CategoriatmpClass = new CategoryClass
			for each y in objTarget.Keys
				if (objTarget(y).getTargetType() = 1) then
					On Error Resume Next
					Set objCategorieXCont = CategoriatmpClass.findCategorieByTargetID(y)
					if not (isNull(objCategorieXCont)) then
						for each z in objCategorieXCont.Keys
							response.write (objCategorieXCont(z).getCatDescrizione() & "<br>")
						next
					end if
					Set objCategorieXCont = nothing
					if(Err.number <>0)then
					end if
				end if									
			next	
			Set CategoriatmpClass = Nothing	
			Set objTarget = nothing
		end if%></td>
		<td class="separator">&nbsp;</td>
		<td><%
		if not(isNull(objFiles)) then
			Dim objFilesInNews
			for each z in objFiles.Keys
				Set objFilesInNews = objFiles(z)
				response.write objFilesInNews.getFileName() & "<br>"
				Set objFilesInNews = nothing	
			next
			Set objFiles = nothing
		end if%></td>
		</tr>
		<tr>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.inserted_date")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.data_pub")%></th>
		<td class="separator">&nbsp;</td>
		<th><%=langEditor.getTranslated("backend.news.view.table.label.data_del")%></th>
		</tr>
		<tr>
		<td><%=dtData_ins%></td>
		<td class="separator">&nbsp;</td>
		<td><%=dtData_pub%></td>
		<td class="separator">&nbsp;</td>
		<td><%=dtData_del%></td>
		</tr>
		
		<tr>
		<th colspan="6"><%=langEditor.getTranslated("backend.contenuti.view.table.label.extra_fields")%></th>
		</tr>
		<tr>
		<td colspan="6">
		<%
		On Error Resume next
		if(hasContentFields) then
			for each k in objListContentField
				Set objField  = objListContentField(k)
				
				labelForm = objField.getDescription()
				if not(langEditor.getTranslated("backend.contenuti.detail.table.label."&labelForm)="") then labelForm = langEditor.getTranslated("backend.contenuti.detail.table.label."&labelForm)
				%>
				<span class="labelForm"><%=labelForm%></span>:&nbsp;<%
					select Case objField.getTypeField()
					Case 3,4,5,6
						Dim valueList
						valueList = ""
						Set objListValues = objContentField.getListContentFieldValues(k)
						for each g in objListValues
							valueList = valueList & Server.HTMLEncode(g) & ","
						next
						
						valueList = Left(valueList,InStrRev(valueList,",",-1,1)-1)						
						response.write(valueList)
						
						Set objListValues = nothing
					Case else						
						response.write(Server.HTMLEncode(objField.getSelValue()))
					end select%><br>
				<%Set objField  = nothing
			next
		end if

		Set objListContentField = nothing
		Set objContentField = nothing

		if(Err.number<>0) then
		'response.write(Err.description)
		end if
		%>
		</td>
		</tr>		
		
		</table>
		<%Set objSelNews = nothing%>
		<br/><input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/contenuti/ListaNews.asp?cssClass=LN"%>';" />
		<br/><br/>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>