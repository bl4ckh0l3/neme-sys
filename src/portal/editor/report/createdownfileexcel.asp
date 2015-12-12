<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/objects/DownloadedFilesClass.asp" -->

<%
Response.Buffer = TRUE 
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "filename=excel_downfile.xls"

Dim objDownFile, objListaDownFile, objTmpDownFile
Set objDownFile = new DownloadedFilesClass
Set objUtente = New UserClass

Dim hasDownFile
hasDownFile = false

on error Resume Next
	Set objListaDownFile = objDownFile.getDownloadedFile()
	
	if(objListaDownFile.Count > 0) then
		hasDownFile = true
	end if
	
if Err.number <> 0 then
end if
%>

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
		<th><%=langEditor.getTranslated("backend.downloaded_file.table.header.fileid")%></th>
		<th><%=langEditor.getTranslated("backend.downloaded_file.table.header.usr")%></th> 
		<th><%=langEditor.getTranslated("backend.downloaded_file.table.header.host")%></th>
		<th><%=langEditor.getTranslated("backend.downloaded_file.table.header.usrinfo")%></th>
		<th><%=langEditor.getTranslated("backend.downloaded_file.table.header.filename")%></th>
		<th><%=langEditor.getTranslated("backend.downloaded_file.table.header.filetype")%></th> 
		<th><%=langEditor.getTranslated("backend.downloaded_file.table.header.filepath")%></th>
		<th><%=langEditor.getTranslated("backend.downloaded_file.table.header.downdate")%></th>
	</tr>
<%
if(hasDownFile) then
	for each x in objListaDownFile%>
		<%Set objTmpDownFile = objListaDownFile(x)%>
		<tr align="left">
			<td><%=objTmpDownFile.getIdFile()%></td>
			<td>
			<%
			if(objTmpDownFile.getIdUser()<>"") then
				on error Resume Next
				Set objTmpUser = objUtente.findUserByIDExt(objTmpDownFile.getIdUser(),false)
				'response.Write(objTmpUser.getCognome() & "&nbsp;&nbsp;" & objTmpUser.getNome())
				response.Write(objTmpUser.getUsername())
				if Err.number <> 0 then
					response.write(Err.description)
				end if	
			end if
			%></td>  
			<td><%=objTmpDownFile.getUserHost()%></td> 
			<td><%=objTmpDownFile.getUserInfo()%></td> 
			<td><%=objTmpDownFile.getFileName()%></td> 
			<td><%=objTmpDownFile.getFileType()%></td> 
			<td><%=objTmpDownFile.getFilePath()%></td> 
			<td><%=objTmpDownFile.getDownloadDate()%></td> 
		</tr>
		<%Set objTmpDownFile = nothing
	next
end if%>

</TABLE>
</body>
</html>
<%
Set objListaDownFile = nothing
Set objUtente = nothing
Set objDownFile = nothing%>