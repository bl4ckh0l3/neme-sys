<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function editCategoria(idCat){
	location.href='<%=Application("baseroot") & "/editor/categorie/InserisciCategoria.asp?cssClass=LCE&id_categoria="%>'+idCat;
}

function deleteCategoria(id_objref, row, refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.categorie.lista.js.alert.delete_category")%>?")){
		//document.form_delete.id_categoria.value = idCat;
		//document.form_delete.submit();
		
		ajaxDeleteItem(id_objref,"category",row, refreshrows);
	}
}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LCE"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
			<div id="ajaxresp" align="center" style="background-color:#FFFF00; border:1px solid #000000; color:#000000; display:none;"></div>		
			<table border="0" cellpadding="0" cellspacing="0" class="principal">
			      <tr> 
				<th colspan="2">&nbsp;</td>
				<th><%=langEditor.getTranslated("backend.categorie.lista.table.header.num_menu")%></th>
				<th><%=langEditor.getTranslated("backend.categorie.lista.table.header.gerarchia")%></th>
				<th><%=langEditor.getTranslated("backend.categorie.lista.table.header.descrizione")%></th>
				<th><%=langEditor.getTranslated("backend.categorie.lista.table.header.visible")%></th>
				<th><%=langEditor.getTranslated("backend.categorie.lista.table.header.type_cat")%></th>
				<th><%=langEditor.getTranslated("backend.categorie.lista.table.header.contiene_news")%></th>
<!--nsys-catlist1-->
				<th><%=langEditor.getTranslated("backend.categorie.lista.table.header.contiene_prod")%></th>
<!---nsys-catlist1-->
			      </tr>
					<%
					Dim hasCategoria
					hasCategoria = false
					on error Resume Next
						Set objListaCategorie = objCategoria.getListaCategorie()
						
						if(objListaCategorie.Count > 0) then
							hasCategoria = true
						end if
						
					if Err.number <> 0 then
						'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
					end if	
					
					if(hasCategoria) then
					
						Dim intCount, tmpObjCat
						intCount = 0
				
				
						'************** codice di prova per paginazione
						
						Dim catsCounter, iIndex, objTmpCats, objTmpCatsKey, FromCat, ToCat, Diff
						iIndex = objListaCategorie.Count
						FromCat = ((numPage * itemsXpage) - itemsXpage)
						Diff = (iIndex - ((numPage * itemsXpage)-1))
						if(Diff < 1) then
							Diff = 1
						end if
						
						ToCat = iIndex - Diff
						
						totPages = iIndex\itemsXpage
						if(totPages < 1) then
							totPages = 1
						elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
							totPages = totPages +1	
						end if		
					
						Dim styleRow, styleRow2
						styleRow2 = "table-list-on"
								
						objTmpCats = objListaCategorie.Items
						objTmpCatsKey=objListaCategorie.Keys		
						for catsCounter = FromCat to ToCat
							styleRow = "table-list-off"
							if(catsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>						
							<form action="<%=Application("baseroot") & "/editor/categorie/CambiaStatoCat.asp"%>" method="post" name="form_change_state_<%=intCount%>">
							<input type="hidden" value="<%=objTmpCatsKey(catsCounter)%>" name="id_cat_to_change">
							<input type="hidden" value="<%=itemsXpage%>" name="items">		
							<input type="hidden" value="<%=numPage%>" name="page">	
							</form>	
							<tr class="<%=styleRow%>" id="tr_delete_list_<%=intCount%>">
							<td align="center" width="25"><a href="javascript:editCategoria(<%=objTmpCatsKey(catsCounter)%>);"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.categorie.lista.table.alt.modify_cat")%>" hspace="2" vspace="0" border="0"></a></td>
							<td align="center" width="25"><a href="javascript:deleteCategoria(<%=objTmpCatsKey(catsCounter)%>,'tr_delete_list_<%=intCount%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.categorie.lista.table.alt.delete_cat")%>" hspace="2" vspace="0" border="0"></a></td>
							<%Set tmpObjCat = objTmpCats(catsCounter)%>
							<td align="center">	
							<div class="ajax" id="view_num_menu_<%=intCount%>" onmouseover="javascript:showHide('view_num_menu_<%=intCount%>','edit_num_menu_<%=intCount%>','num_menu_<%=intCount%>',500, false);"><%=tmpObjCat.getNumMenu()%></div>
							<div class="ajax" id="edit_num_menu_<%=intCount%>"><input type="text" class="formfieldAjaxShort" id="num_menu_<%=intCount%>" name="num_menu" onmouseout="javascript:restoreField('edit_num_menu_<%=intCount%>','view_num_menu_<%=intCount%>','num_menu_<%=intCount%>','category',<%=tmpObjCat.getCatID()%>,1,<%=intCount%>);" value="<%=tmpObjCat.getNumMenu()%>" maxlength="2" onkeypress="javascript:return isInteger(event);"></div>
							<script>
							$("#edit_num_menu_<%=intCount%>").hide();
							</script>
							</td>
							<td width="17%">	
							<div class="ajax" id="view_gerarchia_<%=intCount%>" onmouseover="javascript:showHide('view_gerarchia_<%=intCount%>','edit_gerarchia_<%=intCount%>','gerarchia_<%=intCount%>',500, false);"><%=tmpObjCat.getCatGerarchia()%></div>
							<div class="ajax" id="edit_gerarchia_<%=intCount%>"><input type="text" class="formfieldAjax" id="gerarchia_<%=intCount%>" name="gerarchia" onmouseout="javascript:restoreField('edit_gerarchia_<%=intCount%>','view_gerarchia_<%=intCount%>','gerarchia_<%=intCount%>','category',<%=tmpObjCat.getCatID()%>,1,<%=intCount%>);" value="<%=tmpObjCat.getCatGerarchia()%>" onkeypress="javascript:return isDecimal(event);"></div>
							<script>
							$("#edit_gerarchia_<%=intCount%>").hide();
							</script>
							</td>
							<td><%=tmpObjCat.getCatDescrizione()%></td>
							<td>
							<div class="ajax" id="view_visibile_<%=intCount%>" onmouseover="javascript:showHide('view_visibile_<%=intCount%>','edit_visibile_<%=intCount%>','visibile_<%=intCount%>',500, true);">
							<%
							if (tmpObjCat.isCatVisible() = true) then 
								response.Write(langEditor.getTranslated("backend.commons.yes"))
							else 
								response.Write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</div>
							<div class="ajax" id="edit_visibile_<%=intCount%>">
							<select name="visibile" class="formfieldAjaxSelect" id="visibile_<%=intCount%>" onblur="javascript:updateField('edit_visibile_<%=intCount%>','view_visibile_<%=intCount%>','visibile_<%=intCount%>','category',<%=tmpObjCat.getCatID()%>,2,<%=intCount%>);">
							<OPTION VALUE="0" <%if (tmpObjCat.isCatVisible() = false) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
							<OPTION VALUE="1" <%if (tmpObjCat.isCatVisible() = true) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
							</SELECT>	
							</div>
							<script>
							$("#edit_visibile_<%=intCount%>").hide();
							</script>
							</td>
							<td><%=tmpObjCat.getCatType()%></td>
							<td>
							<div class="ajax" id="view_contiene_news_<%=intCount%>" onmouseover="javascript:showHide('view_contiene_news_<%=intCount%>','edit_contiene_news_<%=intCount%>','contiene_news_<%=intCount%>',500, true);">
							<%if(tmpObjCat.contieneNews()) then
								 response.write(langEditor.getTranslated("backend.commons.yes"))
							else
								response.write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</div>
							<div class="ajax" id="edit_contiene_news_<%=intCount%>">
							<select name="contiene_news" class="formfieldAjaxSelect" id="contiene_news_<%=intCount%>" onblur="javascript:updateField('edit_contiene_news_<%=intCount%>','view_contiene_news_<%=intCount%>','contiene_news_<%=intCount%>','category',<%=tmpObjCat.getCatID()%>,2,<%=intCount%>);">
							<OPTION VALUE="0" <%if (tmpObjCat.contieneNews() = false) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
							<OPTION VALUE="1" <%if (tmpObjCat.contieneNews() = true) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
							</SELECT>	
							</div>
							<script>
							$("#edit_contiene_news_<%=intCount%>").hide();
							</script>
							</td>
<!--nsys-catlist2-->
							<td>
							<div class="ajax" id="view_contiene_prod_<%=intCount%>" onmouseover="javascript:showHide('view_contiene_prod_<%=intCount%>','edit_contiene_prod_<%=intCount%>','contiene_prod_<%=intCount%>',500, true);">
							<%if(tmpObjCat.contieneProd()) then
								 response.write(langEditor.getTranslated("backend.commons.yes"))
							else
								response.write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</div>
							<div class="ajax" id="edit_contiene_prod_<%=intCount%>">
							<select name="contiene_prod" class="formfieldAjaxSelect" id="contiene_prod_<%=intCount%>" onblur="javascript:updateField('edit_contiene_prod_<%=intCount%>','view_contiene_prod_<%=intCount%>','contiene_prod_<%=intCount%>','category',<%=tmpObjCat.getCatID()%>,2,<%=intCount%>);">
							<OPTION VALUE="0" <%if (tmpObjCat.contieneProd() = false) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
							<OPTION VALUE="1" <%if (tmpObjCat.contieneProd() = true) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
							</SELECT>	
							</div>
							<script>
							$("#edit_contiene_prod_<%=intCount%>").hide();
							</script>
							</td>
<!---nsys-catlist2-->
							<%Set tmpObjCat = nothing%>
							</tr>			
							<%intCount = intCount +1
						next
						Set objListaCategorie = nothing		
						%>			  
					  <tr> 
						<form action="<%=Application("baseroot") & "/editor/categorie/ListaCategorie.asp"%>" method="post" name="item_x_page">
<!--nsys-catlist3-->
						<th colspan="9" align="left">
<!---nsys-catlist3-->
						<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
						<%
						'**************** richiamo paginazione
						call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/categorie/ListaCategorie.asp", "&items="&itemsXpage)%>
						</th>
						</form>
					  </tr>
				  <%end if
				  Set objCategoria = Nothing%>
			</table>
			<br/>
			
			<form action="<%=Application("baseroot") & "/editor/categorie/ProcessCategoria.asp"%>" method="post" name="form_delete">
			<input type="hidden" value="" name="id_categoria">
			<input type="hidden" value="del" name="delete_categoria">
			<input type="hidden" value="LCE" name="cssClass">
			</form>
			
			<form action="<%=Application("baseroot") & "/editor/categorie/InserisciCategoria.asp"%>" method="post" name="form_crea">
			<input type="hidden" value="LCE" name="cssClass">
			<input type="hidden" value="-1" name="id_categoria">
			<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.categorie.lista.button.label.inserisci")%>" onclick="javascript:document.form_crea.submit();" />
			</form>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>