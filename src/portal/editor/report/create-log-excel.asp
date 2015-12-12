<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->

<%
Response.Buffer = TRUE 
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "filename=excel_logfile.xls"

Dim objLog, objListaLog, objTmpLog
Dim counter, tmpObjLog, iIndex

Set objLog = New LogClass
Set objListaLog = objLog.getListaLogs(null, null, null)
iIndex = objListaLog.Count
objTmpLog = objListaLog.Items%>

<html>
<head>
<title></title>
<style type="text/css"> 
body {
	background: #FFFFFF;
}
.tdHeaderExcel {
	background-color: #432D30;
	text-align: left;
	color: #FFFFFF;
}
</style>
</head>
<body>
<TABLE BORDER=1>
	<tr class="tdHeaderExcel">
		<td><strong><%=langEditor.getTranslated("backend.logs.include.table.header.msg")%></strong></td>
		<td><strong><%=langEditor.getTranslated("backend.logs.include.table.header.usr")%></strong></td>  
		<td><strong><%=langEditor.getTranslated("backend.logs.include.table.header.type")%></strong></td> 
		<td><strong><%=langEditor.getTranslated("backend.logs.include.table.header.date")%></strong></td> 
	</tr>
<%for counter = 0 to iIndex-1%>
	<%Set tmpObjLog = objTmpLog(counter)%>
	<tr>
		<td><%=tmpObjLog.getLogMsg()%></td>
		<td><%=tmpObjLog.getLogUsr()%></td>  
		<td><%=tmpObjLog.getLogTipo()%></td> 
		<td><%=tmpObjLog.getLogData%></td>
	</tr>
	<%Set tmpObjLog = nothing
next%>

</TABLE>
</body>
</html>
<%
Set objListaLog = nothing
Set objLog = Nothing%>