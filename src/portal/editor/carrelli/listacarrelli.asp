<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/CardClass.asp" -->
<!-- #include virtual="/common/include/Objects/ProductsCardClass.asp" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script>
function deleteCard(id_objref, row,refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.carrello.detail.js.alert.confirm_del_card")%>")){		
		ajaxDeleteItem(id_objref,"shopping_card",row,refreshrows);
	}
}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LCI"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table border="0" cellpadding="0" cellspacing="0" class="principal">
			<tr> 
				<th colspan="2">&nbsp;</th>
				<th><%=langEditor.getTranslated("backend.carrello.lista.table.header.id_carrello")%></th>
				<th><%=langEditor.getTranslated("backend.carrello.lista.table.header.cliente")%></th>
				<th><%=langEditor.getTranslated("backend.carrello.lista.table.header.data_insert")%></th>
			</tr>
			  
			<%
			Dim hasCarrello
			hasCarrello = false
			on error Resume Next
				Set objListaCarrelli = objCarrelli.getListaCarrelli()		
				
				if(objListaCarrelli.Count > 0) then
					hasCarrello = true
				end if
				
			if Err.number <> 0 then
			end if	
			
			if(hasCarrello) then							
				
				intCount = 0										
				iIndex = objListaCarrelli.Count
				
				FromCarrello = ((numPage * itemsXpage) - itemsXpage)
				Diff = (iIndex - ((numPage * itemsXpage)-1))
				if(Diff < 1) then
					Diff = 1
				end if
				
				ToCarrello = iIndex - Diff
				
				totPages = iIndex\itemsXpage
				if(totPages < 1) then
					totPages = 1
				elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
					totPages = totPages +1	
				end if		
						
				objTmpCarrello = objListaCarrelli.Items
				
				Dim objTmpUser, objFilteredCarrello
				Set objUtente = New UserClass					
				
				styleRow2 = "table-list-on"					
						
				for carrelloCounter = FromCarrello to ToCarrello
					styleRow = "table-list-off"
					if(carrelloCounter MOD 2 = 0) then styleRow = styleRow2 end if
					Set objFilteredCarrello = objTmpCarrello(carrelloCounter)
					%>
					<tr class="<%=styleRow%>" id="tr_delete_list_<%=carrelloCounter%>">
					<td align="center" width="25"><a href="<%=Application("baseroot") & "/editor/carrelli/VisualizzaCarrello.asp?cssClass=LCI&id_carrello=" & objFilteredCarrello.getIDCarrello()%>"><img src="<%=Application("baseroot")&"/editor/img/zoom.png"%>" alt="<%=langEditor.getTranslated("backend.carrello.lista.table.alt.view_carrello")%>" hspace="2" vspace="0" border="0"></a></td>
					<td align="center" width="25"><%if(Application("del_carrello_on_exit") = 0) then%><a href="javascript:deleteCard(<%=objFilteredCarrello.getIDCarrello()%>, 'tr_delete_list_<%=carrelloCounter%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.carrello.lista.table.alt.delete_carrello")%>" hspace="2" vspace="0" border="0"></a><%else response.write("&nbsp;") end if%></td>						
					<td><%=objFilteredCarrello.getIDCarrello()%></td>
					<td>
					<%
					on error Resume Next
					Set objTmpUser = objUtente.findUserByIDExt(objFilteredCarrello.getIDUtente(), false)
					response.Write(objTmpUser.getUserName())
					Set objTmpUser = nothing
					
					if Err.number <> 0 then
						response.Write(langEditor.getTranslated("backend.commons.sessione")&": "& objFilteredCarrello.getIDUtente())
					end if	
					%>
					</td>
					<td>
					<%if(DateDiff("d",objFilteredCarrello.getDtaCreazione(),Now()) > Application("day_carrello_is_valid")) then%>
					<span class="carrello_too_old"><%=objFilteredCarrello.getDtaCreazione()%></span>
					<%else%>
					<%=objFilteredCarrello.getDtaCreazione()%>
					<%end if%>
					</td>	
					</tr>				
					<%intCount = intCount +1
					Set objFilteredCarrello = nothing
				next
				Set objUtente = nothing
				Set objTmpCarrello = nothing
				Set objListaCarrelli = nothing
				%>
			  
			  <tr> 
				<form action="<%=Application("baseroot") & "/editor/carrelli/ListaCarrelli.asp"%>" method="post" name="item_x_page">
				<th colspan="5" align="left">
				<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
				<%		
				'**************** richiamo paginazione
				call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/carrelli/ListaCarrelli.asp", "&order_by="&order_carrello_by&"&items="&itemsXpage)%>
				</th>
				</form>			
			</tr>		
			<%end if
			Set objCarrelli = Nothing%>
		</table>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>