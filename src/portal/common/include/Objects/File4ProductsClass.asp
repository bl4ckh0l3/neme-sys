<%
Class File4ProductsClass
	Private id_prodotto
	Private id_attach
	Private file_name
	Private file_type
	Private path
	Private file_dida
	Private file_label
	
	Public Sub setProdottoID(id_prod)
		id_prodotto = id_prod
	End Sub
		
	Public Sub setFileID(id_)
		id_attach = id_
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
	
	
	Public Function getProdottoID()
		getProdottoID = id_prodotto
	End Function
		
	Public Function getFileID()
		getFileID = id_attach
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
	
	
	Public Function getFilePerProdotto(id_prodotto)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		Set getFilePerProdotto = null		
		strSQL = "SELECT * FROM attach_x_prodotti WHERE id_prodotto=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prodotto)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objFiles = new File4ProductsClass
				strID = objRS("id_attach")		
				objFiles.setProdottoID(objRS("id_prodotto"))
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
							
			Set getFilePerProdotto = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function
	
	Public Function getFileByID(id_attach)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objFiles
		getFileByID = null		
		strSQL = "SELECT * FROM attach_x_prodotti WHERE id_attach=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_attach)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objFiles = new File4ProductsClass
			objFiles.setProdottoID(objRS("id_prodotto"))
			objFiles.setFileID(objRS("id_attach"))
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
		strSQL = "SELECT * FROM attach_x_prodotti WHERE filename=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,filename)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objFiles = new File4ProductsClass
			objFiles.setProdottoID(objRS("id_prodotto"))
			objFiles.setFileID(objRS("id_attach"))
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
	
	Public Function getFileByFileNameAndIdProd(idProd, filename)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objFiles
		getFileByFileNameAndIdProd = null		
		strSQL = "SELECT * FROM attach_x_prodotti WHERE id_prodotto=? AND filename=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,filename)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objFiles = new File4ProductsClass
			objFiles.setProdottoID(objRS("id_prodotto"))
			objFiles.setFileID(objRS("id_attach"))
			objFiles.setFileName(objRS("filename"))
			objFiles.setFileType(objRS("content_type"))
			objFiles.setFilePath(objRS("path"))
			objFiles.setFileDida(objRS("file_dida"))
			objFiles.setFileTypeLabel(objRS("file_label"))
							
			Set getFileByFileNameAndIdProd = objFiles
			Set objFiles = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function

	Public Sub insertFileXProdotto(id_prodotto, filename, content_type, path, dida, label, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "INSERT INTO attach_x_prodotti(id_prodotto, filename, content_type, path, file_dida, file_label) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,label)
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

	Public Sub insertFileXProdottoNoTransaction(id_prodotto, filename, content_type, path, dida, label)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		strSQL = "INSERT INTO attach_x_prodotti(id_prodotto, filename, content_type, path, file_dida, file_label) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,label)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyFileXProdotto(id, id_prodotto, filename, content_type, path, dida, label, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE attach_x_prodotti SET "
		strSQL = strSQL & "id_prodotto=?,"
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "path=?,"
		strSQL = strSQL & "file_dida=?,"
		strSQL = strSQL & "file_label=?"
		strSQL = strSQL & " WHERE id_attach=?;" 

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,label)
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
		
	Public Sub modifyFileXProdottoNoTransaction(id, id_prodotto, filename, content_type, path, dida, label)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE attach_x_prodotti SET "
		strSQL = strSQL & "id_prodotto=?,"
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "path=?,"
		strSQL = strSQL & "file_dida=?,"
		strSQL = strSQL & "file_label=?"
		strSQL = strSQL & " WHERE id_attach=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,filename)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,content_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,path)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,dida)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub				
		
	Public Sub deleteFileXProdotto(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM attach_x_prodotti WHERE id_attach=?;"

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
		
	Public Sub deleteFileXProdottoNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn 
		strSQL = "DELETE FROM attach_x_prodotti WHERE id_attach=?;"

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
		strSQL = "SELECT MAX(id_attach) FROM attach_x_prodotti;"

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