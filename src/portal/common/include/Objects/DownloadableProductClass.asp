<%
Class DownloadableProductClass
	Private UploadRequest
	
	Private id
	Private idProd
	Private fileName
	Private filePath
	Private filesize
	Private fileContentType
	Private insertDate
	
	Public Function getID()
		getID = id
	End Function	
	Public Sub setID(strID)
		id = strID
	End Sub
	
	Public Function getIdProd()
		getIdProd = idProd
	End Function	
	Public Sub setIdProd(strIdProd)
		idProd = strIdProd
	End Sub
	
	Public Function getFileName()
		getFileName = fileName
	End Function	
	Public Sub setFileName(strFileName)
		fileName = strFileName
	End Sub
	
	Public Function getFilePath()
		getFilePath = filePath
	End Function	
	Public Sub setFilePath(strFilePath)
		filePath = strFilePath
	End Sub
	
	Public Function getFileSize()
		getFileSize = fileSize
	End Function	
	Public Sub setFileSize(strFileSize)
		fileSize = strFileSize
	End Sub
	
	Public Function getContentType()
		getContentType = fileContentType
	End Function	
	Public Sub setContentType(strContentType)
		fileContentType = strContentType
	End Sub
	
	Public Function getInsertDate()
		getInsertDate = insertDate
	End Function	
	Public Sub setInsertDate(strInsertDate)
		insertDate = strInsertDate
	End Sub
	
	Private Sub Class_Initialize()
		Set UploadRequest = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(UploadRequest) Then
			UploadRequest.RemoveAll()
			Set UploadRequest = Nothing
		End If
	End Sub
	
	Private Sub BuildUploadRequest(RequestBin)	
		Dim PosBeg,PosEnd,boundary,boundaryPos
		PosBeg = 1
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
		boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
		boundaryPos = InstrB(1,RequestBin,boundary)
		
		Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
			Dim UploadControl,Pos,Name,PosFile,PosBound,FileName,ContentType,Value
			Set UploadControl = CreateObject("Scripting.Dictionary")
			Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
			Pos = InstrB(Pos,RequestBin,getByteString("name="))
			PosBeg = Pos+6
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
			PosBound = InstrB(PosEnd,RequestBin,boundary)
			If  PosFile<>0 AND (PosFile<PosBound) Then
				PosBeg = PosFile + 10
				PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
				FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "FileName", FileName
				Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
				PosBeg = Pos+14
				PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
				ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "ContentType",ContentType
				PosBeg = PosEnd+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			Else
				Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
				PosBeg = Pos+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			End If
			UploadControl.Add "Value" , Value	
			UploadRequest.Add name, UploadControl	
			BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
		Loop
	End Sub
	
	Function getByteString(StringStr)
		Dim char
		For i = 1 to Len(StringStr)
			char = Mid(StringStr,i,1)
			getByteString = getByteString & chrB(AscB(char))
		Next
	End Function
	
	Function getString(StringBin)
		getString =""
		Dim intCount
		For intCount = 1 to LenB(StringBin)
			getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
		Next
	End Function
	
	Public Sub saveDownloadProd(id_prod, fieldNamePrefix, maxItem, varArrayBinRequest)		
		'Create FileSytemObject Component
		Dim ScriptObject
		Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")	
		
		Dim dirUpload, RequestBin
		'attenzione, unica riga da modificare
		'inserire il percorso della sotto-cartella in public, ESISTENTE, nella quale verranno inseriti i files
		'-----------------------------------------------------
		dirUpload = Server.MapPath(Application("baseroot")&Application("dir_down_prod"))
		dirUpload = dirUpload & "\" & id_prod &"\"
		if not(ScriptObject.FolderExists(dirUpload)) then
			call ScriptObject.CreateFolder(dirUpload)	
		end if
	
		'------------------------------------------------------
		'nota, se vuoi fare upload in cartella nella quale i files siano raggiungibili solo via FTP (per massima sicurezza)
		'puoi cambiare il percorso, ad esempio con "/mdb-database/nomecartella/"
		'fine modifica linkbc 07/07/2008	
		
		On Error Resume Next
			
		'RequestBin = Request.BinaryRead(Request.TotalBytes)
		RequestBin = varArrayBinRequest
		
		Dim filepathname, filename, value, pathEnd
		call BuildUploadRequest(RequestBin)
		
		Dim y, MyFile
		for y = 1 to maxItem	
			if not(isEmpty(UploadRequest(CStr(fieldNamePrefix&y))))then
				'contentType = UploadRequest.Item("blob").Item("ContentType")
				filepathname = UploadRequest(CStr(fieldNamePrefix&y)).Item("FileName")
				filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
				value = UploadRequest(CStr(fieldNamePrefix&y)).Item("Value")		
				pathEnd = dirUpload & filename		
				Set MyFile = ScriptObject.CreateTextFile(pathEnd, true)
				For i = 1 to LenB(value)
					MyFile.Write chr(AscB(MidB(value,i,1)))
				Next
				MyFile.Close
			end if
		next
		if (Err.number <> 0) then
			'response.Write(Err.description)
		end if	
		
		Set ScriptObject = nothing
	end Sub
	
	Public Function getFilePerProdotto(id_prodotto)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		Set getFilePerProdotto = null		
		strSQL = "SELECT * FROM downloadable_products WHERE id_product=?;"
		
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
				Set objDownProd = new DownloadableProductClass
				strID = objRS("id")		
				objDownProd.setID(strID)
				objDownProd.setIdProd(objRS("id_product"))
				objDownProd.setFileName(objRS("filename"))
				objDownProd.setFilePath(objRS("path"))
				objDownProd.setFileSize(objRS("file_size"))
				objDownProd.setContentType(objRS("content_type"))
				objDownProd.setInsertDate(objRS("insert_date"))			
			
				objDict.add strID, objDownProd
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
	
	Public Function getFileByID(id_file)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDownProd
		getFileByID = null		
		strSQL = "SELECT * FROM downloadable_products WHERE id=?;"
				
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
			Set objDownProd = new DownloadableProductClass
			objDownProd.setID(objRS("id"))
			objDownProd.setIdProd(objRS("id_product"))
			objDownProd.setFileName(objRS("filename"))
			objDownProd.setFilePath(objRS("path"))
			objDownProd.setFileSize(objRS("file_size"))
			objDownProd.setContentType(objRS("content_type"))
			objDownProd.setInsertDate(objRS("insert_date"))
							
			Set getFileByID = objDownProd
			Set objDownProd = nothing				
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
		Dim objDB, strSQL, objRS, objConn, objDownProd
		getFileByFileName = null		
		strSQL = "SELECT * FROM downloadable_products WHERE filename=?;"
				
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
			Set objDownProd = new DownloadableProductClass
			objDownProd.setID(objRS("id"))
			objDownProd.setIdProd(objRS("id_product"))
			objDownProd.setFileName(objRS("filename"))
			objDownProd.setFilePath(objRS("path"))
			objDownProd.setFileSize(objRS("file_size"))
			objDownProd.setContentType(objRS("content_type"))
			objDownProd.setInsertDate(objRS("insert_date"))
							
			Set getFileByFileName = objDownProd
			Set objDownProd = nothing				
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
		Dim objDB, strSQL, objRS, objConn, objDownProd
		getFileByFileNameAndIdProd = null		
		strSQL = "SELECT * FROM downloadable_products WHERE id_product=? AND filename=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filename)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objDownProd = new DownloadableProductClass
			objDownProd.setID(objRS("id"))
			objDownProd.setIdProd(objRS("id_product"))
			objDownProd.setFileName(objRS("filename"))
			objDownProd.setFilePath(objRS("path"))
			objDownProd.setFileSize(objRS("file_size"))
			objDownProd.setContentType(objRS("content_type"))
			objDownProd.setInsertDate(objRS("insert_date"))
							
			Set getFileByFileNameAndIdProd = objDownProd
			Set objDownProd = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function

	'********************************************* METODI DAO *********************************************
	Public Function insertDownProd(idDownProd, strDownFileName, strDownFilePath, strDownContentType, strDownFilesize, objConn)
		on error resume next
		insertDownProd = -1
		
		Dim strSQL, objRS, dtData_ins
		
		dtData_ins = convertDate(now())
		
		strSQL = "INSERT INTO downloadable_products(id_product, filename, path, content_type, file_size, insert_date) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idDownProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDownFileName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDownFilePath)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strDownContentType)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strDownFilesize)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(downloadable_products.id) as id FROM downloadable_products")
		if not (objRS.EOF) then
			'objRS.MoveFirst()
			insertDownProd = objRS("id")	
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

	Public Function insertDownProdNoTransaction(idDownProd, strDownFileName, strDownFilePath, strDownContentType, strDownFilesize)
		on error resume next
		insertDownProdNoTransaction = -1
		
		Dim objDB, strSQL, objRS, objConn, dtData_ins
		
		dtData_ins = convertDate(now())
		
		strSQL = "INSERT INTO downloadable_products(id_product, filename, path, content_type, file_size, insert_date) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idDownProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDownFileName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDownFilePath)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strDownContentType)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strDownFilesize)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(downloadable_products.id) as id FROM downloadable_products")
		if not (objRS.EOF) then
			'objRS.MoveFirst()
			insertDownProdNoTransaction = objRS("id")	
		end if	
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyDownProd(idThis, idDownProd, strDownFileName, strDownFilePath, strDownContentType, strDownFilesize, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		dtData_ins = convertDate(now())
		
		strSQL = "UPDATE downloadable_products SET "
		strSQL = strSQL & "id_product=?,"
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "path=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "file_size=?,"
		strSQL = strSQL & "insert_date=?"
		strSQL = strSQL & " WHERE id_prodotto=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idDownProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDownFileName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDownFilePath)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strDownContentType)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strDownFilesize)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idThis)
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
		
	Public Sub modifyDownProdNoTransaction(idThis, idDownProd, strDownFileName, strDownFilePath, strDownContentType, strDownFilesize)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		dtData_ins = convertDate(now())
		
		strSQL = "UPDATE downloadable_products SET "
		strSQL = strSQL & "id_product=?,"
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "path=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "file_size=?,"
		strSQL = strSQL & "insert_date=?"
		strSQL = strSQL & " WHERE id_prodotto=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idDownProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDownFileName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDownFilePath)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strDownContentType)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strDownFilesize)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idThis)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteDownProd(idThis, objConn)
		on error resume next
		Dim objDB, strSQLDelProdotto, objRS
		strSQLDelProdotto = "DELETE FROM downloadable_products WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDelProdotto
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idThis)
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
		
	Public Sub deleteDownProdNoTransaction(idThis)
		on error resume next
		Dim objDB, strSQLDelProdotto, objRS, objConn
		strSQLDelProdotto = "DELETE FROM downloadable_products WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDelProdotto
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idThis)
		objCommand.Execute()
		Set objCommand = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	

	
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

End Class
%>