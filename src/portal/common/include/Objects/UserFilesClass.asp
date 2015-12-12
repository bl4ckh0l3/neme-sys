<%
Class UserFilesClass
	Private id_file
	Private id_user
	Private file_name
	Private file_type
	Private path
	Private file_dida
	Private file_label
	Private data_ins
	
	Public Sub setFileID(id_)
		id_file = id_
	End Sub	
	
	Public Sub setUserID(idUser)
		id_user = idUser
	End Sub	
	
	Public Sub setFileName(name_)
		file_name = name_
	End Sub

	Public Sub setFileType(type_)
		file_type = type_
	End Sub

	Public Sub setFilePath(path_)
		path = path_
	End Sub

	Public Sub setFileDida(dida_)
		file_dida = dida_
	End Sub

	Public Sub setFileTypeLabel(filelabel_)
		file_label = filelabel_
	End Sub

	Public Sub setDataIns(dataIns)
		data_ins = dataIns
	End Sub
	
	
	Public Function getFileID()
		getFileID = id_file
	End Function	
	
	Public Function getUserID()
		getUserID = id_user
	End Function	
	
	Public Function getFileName()
		getFileName = file_name
	End Function

	Public Function getFileType()
		getFileType = file_type
	End Function

	Public Function getFilePath()
		getFilePath = path
	End Function

	Public Function getFileDida()
		getFileDida = file_dida
	End Function

	Public Function getFileTypeLabel()
		getFileTypeLabel = file_label
	End Function

	Public Function getDataIns()
		getDataIns = data_ins
	End Function
	
	
	Public Function getFiles4User(id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getFiles4User = null		
		strSQL = "SELECT * FROM user_files WHERE id_user=? ORDER BY dta_ins DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objFiles = new UserFilesClass
				strID = objRS("id")
				objFiles.setFileID(strID)
				objFiles.setUserID(objRS("id_user"))
				objFiles.setFileName(objRS("filename"))
				objFiles.setFileType(objRS("content_type"))
				objFiles.setFilePath(objRS("path"))
				objFiles.setFileDida(objRS("file_dida"))
				objFiles.setFileTypeLabel(objRS("file_label"))		
				objFiles.setDataIns(objRS("dta_ins"))			
						
				objDict.add strID, objFiles
				Set objFiles = nothing
				objRS.moveNext()
			loop
							
			Set getFiles4User = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function
	
	Public Function getFilesByID(id_file)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objFiles
		getFilesByID = null		
		strSQL = "SELECT * FROM user_files WHERE id=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_file)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objFiles = new UserFilesClass
			objFiles.setFileID(objRS("id"))
			objFiles.setUserID(objRS("id_user"))
			objFiles.setFileName(objRS("filename"))
			objFiles.setFileType(objRS("content_type"))
			objFiles.setFilePath(objRS("path"))
			objFiles.setFileDida(objRS("file_dida"))
			objFiles.setFileTypeLabel(objRS("file_label"))	
			objFiles.setDataIns(objRS("dta_ins"))	
							
			Set getFilesByID = objFiles
			Set objFiles = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function	
	
	Public Function getFilesByFileName(filename)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objFiles
		getFilesByFileName = null		
		strSQL = "SELECT * FROM user_files WHERE filename=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objFiles = new UserFilesClass
			objFiles.setFileID(objRS("id"))			
			objFiles.setUserID(objRS("id_user"))
			objFiles.setFileName(objRS("filename"))
			objFiles.setFileType(objRS("content_type"))
			objFiles.setFilePath(objRS("path"))
			objFiles.setFileDida(objRS("file_dida"))
			objFiles.setFileTypeLabel(objRS("file_label"))	
			objFiles.setDataIns(objRS("dta_ins"))	
							
			Set getFilesByFileName = objFiles
			Set objFiles = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function		
	
	Public Function getFilesByFileNameAndIdUser(idUser, filename)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objFiles
		getFilesByFileNameAndIdUser = null		
		strSQL = "SELECT * FROM user_files WHERE filename=? AND id_user=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objFiles = new UserFilesClass
			objFiles.setFileID(objRS("id"))
			objFiles.setUserID(objRS("id_user"))
			objFiles.setFileName(objRS("filename"))
			objFiles.setFileType(objRS("content_type"))
			objFiles.setFilePath(objRS("path"))
			objFiles.setFileDida(objRS("file_dida"))
			objFiles.setFileTypeLabel(objRS("file_label"))	
			objFiles.setDataIns(objRS("dta_ins"))	
							
			Set getFilesByFileNameAndIdUser = objFiles
			Set objFiles = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function	
			
	Public Function insertFiles(id_user, filename, content_type, path, dida, label, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, id_new_file, dtData_ins
		
		insertFiles = -1
		
		dtData_ins = now()
		
		if (Application("dbType") = 1) then
			dtData_ins = convertDate(dtData_ins)
		end if	
		
		strSQL = "INSERT INTO user_files(id_user, filename, content_type, path, file_dida, file_label, dta_ins) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Execute()
		Set objCommand = Nothing
	
		Set objRS = objConn.Execute("SELECT max(user_files.id) as id FROM user_files")
		if not (objRS.EOF) then
			insertFiles = objRS("id")	
		end if		
		Set objRS = Nothing	
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
			
	Public Sub insertFilesNoTransaction(id_user, filename, content_type, path, dida, label)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, id_new_file, dtData_ins

		dtData_ins = now()
		
		if (Application("dbType") = 1) then
			dtData_ins = convertDate(dtData_ins)
		end if
		
		strSQL = "INSERT INTO user_files(id_user, filename, content_type, path, file_dida, file_label, dta_ins) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?);"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyFiles(id, id_user, filename, content_type, path, dida, label, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, dtData_ins

		dtData_ins = now()
		
		if (Application("dbType") = 1) then
			dtData_ins = convertDate(dtData_ins)
		end if
		
		strSQL = "UPDATE user_files SET "
		strSQL = strSQL & "id_user=?,"
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "path=?,"
		strSQL = strSQL & "file_dida=?,"
		strSQL = strSQL & "file_label=?,"
		strSQL = strSQL & "dta_ins=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing	
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyFilesNoTransaction(id, id_user, filename, content_type, path, dida, label)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, dtData_ins

		dtData_ins = now()
		
		if (Application("dbType") = 1) then
			dtData_ins = convertDate(dtData_ins)
		end if
		
		strSQL = "UPDATE user_files SET "
		strSQL = strSQL & "id_user=?,"
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "path=?,"
		strSQL = strSQL & "file_dida=?,"
		strSQL = strSQL & "file_label=?,"
		strSQL = strSQL & "dta_ins=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteFiles(id, idUser, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM user_files WHERE id=? AND id_user=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Execute()
		Set objCommand = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteFilesNoTransaction(id, idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn 
		strSQL = "DELETE FROM user_files WHERE id=? AND id_user=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Function getMaxIDFiles()
		on error resume next
		
		getMaxIDFile = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT max(id) as id FROM user_files;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDFile = objRS("id")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function convertDate(dateToConvert)
		Dim DD, MM, YY, HH, MIN, SS
		
		convertDate = null
		
		DD = DatePart("d", dateToConvert)
		MM = DatePart("m", dateToConvert)
		YY = DatePart("yyyy", dateToConvert)
		HH = DatePart("h", dateToConvert)
		MIN = DatePart("n", dateToConvert)
		SS = DatePart("s", dateToConvert)
		
		convertDate = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS		
	End Function
	
	Public Function getListaFilesLabel()
		Set getListaFilesLabel = Server.CreateObject("Scripting.Dictionary")
		getListaFilesLabel.add "1", "img small"
		getListaFilesLabel.add "2", "img big"
		getListaFilesLabel.add "3", "pdf"
		getListaFilesLabel.add "4", "audio-video"
		getListaFilesLabel.add "5", "others..."
		getListaFilesLabel.add "6", "img medium"
	End Function										
End Class
%>