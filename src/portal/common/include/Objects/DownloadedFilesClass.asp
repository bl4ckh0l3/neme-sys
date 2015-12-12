<%
Class DownloadedFilesClass
	Private id
	Private idFile
	Private idUser
	Private userHost
	Private userInfo
	Private fileName
	Private fileType
	Private path
	Private downloadDate

	Public Function getID()
		getID = id
	End Function	
	Public Sub setID(strID)
		id = strID
	End Sub
	
	Public Function getIdFile()
		getIdFile = idFile
	End Function	
	Public Sub setIdFile(strIdFile)
		idFile = strIdFile
	End Sub
	
	Public Function getIdUser()
		getIdUser = idUser
	End Function	
	Public Sub setIdUser(strIdUser)
		idUser = strIdUser
	End Sub	
	
	Public Function getUserHost()
		getUserHost = userHost
	End Function	
	Public Sub setUserHost(strUserHost)
		userHost = strUserHost
	End Sub		
	
	Public Function getUserInfo()
		getUserInfo = userInfo
	End Function	
	Public Sub setUserInfo(strUserInfo)
		userInfo = strUserInfo
	End Sub		
	
	Public Function getFileName()
		getFileName = fileName
	End Function	
	Public Sub setFileName(strFilename)
		fileName = strFilename
	End Sub

	Public Function getFileType()
		getFileType = fileType
	End Function
	Public Sub setFileType(strFileType)
		fileType = strFileType
	End Sub

	Public Function getFilePath()
		getFilePath = path
	End Function
	Public Sub setFilePath(strPath)
		path = strPath
	End Sub
	
	Public Function getDownloadDate()
		getDownloadDate = downloadDate
	End Function	
	Public Sub setDownloadDate(strDownloadDate)
		downloadDate = strDownloadDate
	End Sub	
	
	Public Function getDownloadedFile()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, strID
		Set getDownloadedFile = null		
		strSQL = "SELECT * FROM downloaded_files ORDER BY download_date DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objDownProd = new DownloadedFilesClass
				strID = objRS("id")		
				objDownProd.setID(strID)
				objDownProd.setIdFile(objRS("id_file"))
				objDownProd.setIdUser(objRS("id_user"))
				objDownProd.setUserHost(objRS("user_host"))
				objDownProd.setUserInfo(objRS("user_info"))
				objDownProd.setFileName(objRS("filename"))
				objDownProd.setFileType(objRS("content_type"))
				objDownProd.setFilePath(objRS("path"))			
				objDownProd.setDownloadDate(objRS("download_date"))	
			
				objDict.add strID, objDownProd
				Set objFiles = nothing
				objRS.moveNext()
			loop
							
			Set getDownloadedFile = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function		
	
	Public Function getDownloadedFileByIdFile(id_file)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, strID
		Set getDownloadedFileByIdFile = null		
		strSQL = "SELECT * FROM downloaded_files WHERE id_file=? ORDER BY download_date DESC;"
		
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
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objDownProd = new DownloadedFilesClass
				strID = objRS("id")		
				objDownProd.setID(strID)
				objDownProd.setIdFile(objRS("id_file"))
				objDownProd.setIdUser(objRS("id_user"))
				objDownProd.setUserHost(objRS("user_host"))
				objDownProd.setUserInfo(objRS("user_info"))
				objDownProd.setFileName(objRS("filename"))
				objDownProd.setFileType(objRS("content_type"))
				objDownProd.setFilePath(objRS("path"))			
				objDownProd.setDownloadDate(objRS("download_date"))	
			
				objDict.add strID, objDownProd
				Set objFiles = nothing
				objRS.moveNext()
			loop
							
			Set getDownloadedFileByIdFile = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function	
	
	
	Public Function getDownloadedFileByID(id_file)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDownProd, strID
		getDownloadedFileByID = null		
		strSQL = "SELECT * FROM downloaded_files WHERE id=?;"
				
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
			Set objDownProd = new DownloadedFilesClass
			strID = objRS("id")		
			objDownProd.setID(strID)
			objDownProd.setIdFile(objRS("id_file"))
			objDownProd.setIdUser(objRS("id_user"))
			objDownProd.setUserHost(objRS("user_host"))
			objDownProd.setUserInfo(objRS("user_info"))
			objDownProd.setFileName(objRS("filename"))
			objDownProd.setFileType(objRS("content_type"))
			objDownProd.setFilePath(objRS("path"))			
			objDownProd.setDownloadDate(objRS("download_date"))
							
			Set getDownloadedFileByID = objDownProd
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
	Public Function insertDownFile(IdFile, IdUser, userHost, userInfo, fileName, fileType, filePath, DownloadDate, objConn)
		on error resume next
		insertDownFile = -1
		
		Dim strSQL, objRS, dtData_ins, dtDownloadDate

		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			dtDownloadDate = convertDate(DownloadDate)
		else	
			'controllo il tipo di database in uso
			if (Application("dbType") = 0) then
				dtDownloadDate = "null"
			else		
				dtDownloadDate = "'0000-00-00 00:00:00'"
			end if			
		end if
		
		strSQL = "INSERT INTO downloaded_files(id_file, id_user, user_host, user_info, filename, content_type, path, download_date) VALUES("
		strSQL = strSQL & "?,"
		if(isNull(IdUser) OR IdUser = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ",?,?,?,?,?,"
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			strSQL = strSQL & "?);"
		else
			strSQL = strSQL & dtDownloadDate&");"
		end if
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdFile)
		if not isNull(IdUser) AND not(IdUser = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdUser)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,userHost)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,userInfo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,fileName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,fileType)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filePath)
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtDownloadDate)
		end if
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(downloaded_files.id) as id FROM downloaded_files")
		if not (objRS.EOF) then
			'objRS.MoveFirst()
			insertDownFile = objRS("id")	
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

	Public Function insertDownFileNoTransaction(IdFile, IdUser, userHost, userInfo, fileName, fileType, filePath, DownloadDate)
		on error resume next
		insertDownFileNoTransaction = -1
		
		Dim objDB, strSQL, objRS, objConn, dtData_ins, dtDownloadDate

		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			dtDownloadDate = convertDate(DownloadDate)
		else	
			'controllo il tipo di database in uso
			if (Application("dbType") = 0) then
				dtDownloadDate = "null"
			else		
				dtDownloadDate = "'0000-00-00 00:00:00'"
			end if			
		end if
		
		strSQL = "INSERT INTO downloaded_files(id_file, id_user, user_host, user_info, filename, content_type, path, download_date) VALUES("
		strSQL = strSQL & "?,"
		if(isNull(IdUser) OR IdUser = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ",?,?,?,?,?,"
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			strSQL = strSQL & "?);"
		else
			strSQL = strSQL & dtDownloadDate&");"
		end if

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdFile)
		if not isNull(IdUser) AND not(IdUser = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdUser)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,userHost)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,userInfo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,fileName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,fileType)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filePath)
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtDownloadDate)
		end if
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(downloaded_files.id) as id FROM downloaded_files")
		if not (objRS.EOF) then
			'objRS.MoveFirst()
			insertDownFileNoTransaction = objRS("id")	
		end if	
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyDownFile(idThis, IdFile, IdUser, userHost, userInfo, fileName, fileType, filePath, DownloadDate, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, dtDownloadDate

		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			dtDownloadDate = convertDate(DownloadDate)
		else	
			'controllo il tipo di database in uso
			if (Application("dbType") = 0) then
				dtDownloadDate = "null"
			else		
				dtDownloadDate = "'0000-00-00 00:00:00'"
			end if			
		end if
				
		strSQL = "UPDATE downloaded_files SET "
		strSQL = strSQL & "id_file=?,"
		if(isNull(IdUser) OR IdUser = "") then
			strSQL = strSQL & "id_user=NULL,"
		else
			strSQL = strSQL & "id_user=?,"			
		end if
		strSQL = strSQL & "user_host=?,"
		strSQL = strSQL & "user_info=?,"
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "path=?,"
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			strSQL = strSQL & "download_date=?"
		else
			strSQL = strSQL & "download_date="&dtDownloadDate
		end if
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdFile)
		if not isNull(IdUser) AND not(IdUser = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdUser)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,userHost)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,userInfo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,fileName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,fileType)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filePath)
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtDownloadDate)
		end if
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
		
	Public Sub modifyDownFileNoTransaction(idThis, IdFile, IdUser, userHost, userInfo, fileName, fileType, filePath, DownloadDate)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, dtDownloadDate

		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			dtDownloadDate = convertDate(DownloadDate)
		else	
			'controllo il tipo di database in uso
			if (Application("dbType") = 0) then
				dtDownloadDate = "null"
			else		
				dtDownloadDate = "'0000-00-00 00:00:00'"
			end if			
		end if

		strSQL = "UPDATE downloaded_files SET "
		strSQL = strSQL & "id_file=?,"
		if(isNull(IdUser) OR IdUser = "") then
			strSQL = strSQL & "id_user=NULL,"
		else
			strSQL = strSQL & "id_user=?,"			
		end if
		strSQL = strSQL & "user_host=?,"
		strSQL = strSQL & "user_info=?,"
		strSQL = strSQL & "filename=?,"
		strSQL = strSQL & "content_type=?,"
		strSQL = strSQL & "path=?,"
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			strSQL = strSQL & "download_date=?"
		else
			strSQL = strSQL & "download_date="&dtDownloadDate
		end if
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdFile)
		if not isNull(IdUser) AND not(IdUser = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdUser)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,userHost)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,userInfo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,fileName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,fileType)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,filePath)
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtDownloadDate)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idThis)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteDownFile(idThis, objConn)
		on error resume next
		Dim objDB, strSQLDelProdotto, objRS
		strSQLDelProdotto = "DELETE FROM downloaded_files WHERE id=?;"

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
		
	Public Sub deleteDownFileNoTransaction(idThis)
		on error resume next
		Dim objDB, strSQLDelProdotto, objRS, objConn
		strSQLDelProdotto = "DELETE FROM downloaded_files WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDelProdotto
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idThis)
		objCommand.Execute()
		Set objCommand = Nothings	
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