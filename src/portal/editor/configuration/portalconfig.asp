<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ConfigClass.asp" -->
<!-- #include file="include/init.asp" -->
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
		<%cssClass="CP"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
			<table border="0" cellpadding="0" cellspacing="0" class="principal">
				<tr>
					<th>&nbsp;</th>
					<th>&nbsp;</th>
					<th><%=langEditor.getTranslated("backend.config.lista.table.header.nome_variabile")%></th>
					<th><%=langEditor.getTranslated("backend.config.lista.table.header.descrizione")%></th>
					<th><%=langEditor.getTranslated("backend.config.lista.table.header.value")%></th>
				</tr>
				<%
				Dim intCount, objListaConfig, objTmpConfig, styleRow, styleRow2
				intCount = 0
				Set objListaConfig = objConfig.getListaConfig()
				
				styleRow2 = "table-list-on"

				for each y in objListaConfig.Keys	
					Set objTmpConfig = objListaConfig(y)
					styleRow = "table-list-off"
					if(intCount MOD 2 = 0) then styleRow = styleRow2 end if
					%>

					<form action="<%=Application("baseroot") & "/editor/configuration/PortalConfig.asp"%>" method="post" name="form_lista_<%=intCount%>">
					<input type="hidden" value="<%=objTmpConfig.getKey()%>" name="key">					
					<tr class="<%=styleRow%>">
					<td><%if(objTmpConfig.getAlert() = "1") then%><img src=<%=Application("baseroot")&"/common/img/ico_alert.gif"%> vspace="2" hspace="2" border="0" align="middle" alt="<%=langEditor.getTranslated("backend.config.lista.table.alt.dont_doit")%>"><%else response.write("&nbsp;") end if%></td>
					<td align="center"><!--nsys-democonf1--><a href="javascript:document.form_lista_<%=intCount%>.submit();"><!---nsys-democonf1--><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.config.lista.table.alt.modify_config")%>" hspace="2" vspace="0" border="0"></a></td>
					<td><span class="labelForm"><%=objTmpConfig.getKey()%></span></td>
					<td width="400"><%=langEditor.getTranslated(objTmpConfig.getDescrizione())%></td>
					<td><input type="text" name="value" value="<%=objTmpConfig.getValue()%>" class="formFieldTXT"></td>
					</tr>					
					</form>					
					<%
					Set objTmpConfig = nothing
					intCount = intCount +1
				next
				Set objListaConfig = nothing
				%>			
				<tr>
					<th colspan="5">&nbsp;</th>
				</tr>
			</table>
			<br/>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>