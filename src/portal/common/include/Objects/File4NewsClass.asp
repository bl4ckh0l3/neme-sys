<%
Class File4NewsClass
	Private id_file
	Private file_name
	Private file_type
	Private path
	Private file_dida
	Private file_label
	
	Public Sub setFileID(id_)
		id_file = id_
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
	
	
	Public Function getFileID()
		getFileID = id_file
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
	
	
	Public Function getFilePerNews(id_news)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getFilePerNews = null		
		strSQL = "SELECT uploaded_files.* FROM uploaded_files INNER JOIN file_x_news ON uploaded_files.id =file_x_news.id_file WHERE file_x_news.id_news=? ORDER BY uploaded_files.filename;"
	
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_news)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objFiles = new File4NewsClass
				strID = objRS("id")
				objFiles.setFileID(strID)
				objFiles.setFileName(objRS("filename"))
				objFiles.setFileType(objRS("content_type"))
				objFiles.setFilePath(objRS("path"))
				objFiles.setFileDida(objRS("file_dida"))
				objFiles.setFileTypeLabel(objRS("file_label"))			
						
				objDict.add strID, objFiles
				Set objFiles = nothing
				objRS.moveNext()
			loop
							
			Set getFilePerNews = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function
	
	Public Function getFileByID(id_file)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objFiles
		getFileByID = null		
		strSQL = "SELECT * FROM uploaded_files WHERE id=?;"
				
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
			Set objFiles = new File4NewsClass
			objFiles.setFileID(objRS("id"))
			objFiles.setFileName(objRS("filename"))
			objFiles.setFileType(objRS("content_type"))
			objFiles.setFilePath(objRS("path"))
			objFiles.setFileDida(objRS("file_dida"))
			objFiles.setFileTypeLabel(objRS("file_label"))
							
			Set getFileByID = objFiles
			Set objFiles = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function	
	
	Public Function getFileByFileName(filename)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objFiles
		getFileByFileName = null		
		strSQL = "SELECT * FROM uploaded_files WHERE filename=?;"
				
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
			Set objFiles = new File4NewsClass
			objFiles.setFileID(objRS("id"))
			objFiles.setFileName(objRS("filename"))
			objFiles.setFileType(objRS("content_type"))
			objFiles.setFilePath(objRS("path"))
			objFiles.setFileDida(objRS("file_dida"))
			objFiles.setFileTypeLabel(objRS("file_label"))
							
			Set getFileByFileName = objFiles
			Set objFiles = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function		
	
	Public Function getFileByFileNameAndIdNews(idNews, filename)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objFiles
		getFileByFileNameAndIdNews = null		
		strSQL = "SELECT uploaded_files.* FROM uploaded_files INNER JOIN file_x_news ON uploaded_files.id =file_x_news.id_file WHERE file_x_news.id_news=? AND filename=? ORDER BY uploaded_files.filename;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idNews)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Set objFiles = new File4NewsClass
			objFiles.setFileID(objRS("id"))
			objFiles.setFileName(objRS("filename"))
			objFiles.setFileType(objRS("content_type"))
			objFiles.setFilePath(objRS("path"))
			objFiles.setFileDida(objRS("file_dida"))
			objFiles.setFileTypeLabel(objRS("file_label"))
							
			Set getFileByFileNameAndIdNews = objFiles
			Set objFiles = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function	
			
	Public Function insertFile(filename, content_type, path, dida, label, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, id_new_file
		
		insertFile = -1
		
		strSQL = "INSERT INTO uploaded_files(filename, content_type, path, file_dida, file_label) VALUES("
		strSQL = strSQL & "?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(uploaded_files.id) as id FROM uploaded_files")
		if not (objRS.EOF) then
			insertFile = objRS("id")	
		end if		
		Set objRS = Nothing	
		Set objCommand = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
			
	Public Sub insertFileNoTransaction(filename, content_type, path, dida, label)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, id_new_file
		
		strSQL = "INSERT INTO uploaded_files(filename, content_type, path, file_dida, file_label) VALUES("
		strSQL = strSQL & "?,?,?,?,?);"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyFile(id, filename, content_type, path, dida, label, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE uploaded_files SET "
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "path=?,"
		strSQL = strSQL & "file_dida=?,"
		strSQL = strSQL & "file_label=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
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
		
	Public Sub modifyFileNoTransaction(id, filename, content_type, path, dida, label)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE uploaded_files SET "
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "path=?,"
		strSQL = strSQL & "file_dida=?,"
		strSQL = strSQL & "file_label=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteFile(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM uploaded_files WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
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
		
	Public Sub deleteFileNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn 
		strSQL = "DELETE FROM uploaded_files WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub insertFileXNews(id_news, id_file, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "INSERT INTO file_x_news(id_news, id_file) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_file)
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

	Public Sub insertFileXNewsNoTransaction(id_news, id_file)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		strSQL = "INSERT INTO file_x_news(id_news, id_file) VALUES("
		strSQL = strSQL & "?,?);"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_file)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteFileXNews(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM file_x_news WHERE id_file=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
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
		
	Public Sub deleteFileXNewsNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM file_x_news WHERE id_file=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Function getMaxIDFile()
		on error resume next
		
		getMaxIDFile = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT max(uploaded_files.id) as id FROM uploaded_files;"

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
	
	Public Function getListaFileLabel()
		Set getListaFileLabel = Server.CreateObject("Scripting.Dictionary")
		getListaFileLabel.add "1", "img small"
		getListaFileLabel.add "2", "img big"
		getListaFileLabel.add "3", "pdf"
		getListaFileLabel.add "4", "audio-video"
		getListaFileLabel.add "5", "others..."
		getListaFileLabel.add "6", "img medium"
		getListaFileLabel.add "7", "img carrello"
		getListaFileLabel.add "8", "file protected"
	End Function										
End Class
%>