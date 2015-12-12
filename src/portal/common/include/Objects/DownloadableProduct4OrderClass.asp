<%
Class DownloadableProduct4OrderClass
	Private id
	Private idOrder
	Private idProd
	Private idDownProd
	Private iduser
	Private active
	Private maxNumDownload
	Private insertDate
	Private expireDate
	Private downloadCounter
	Private downloadDate

	Public Function getID()
		getID = id
	End Function	
	Public Sub setID(strID)
		id = strID
	End Sub
	
	Public Function getIdOrder()
		getIdOrder = idOrder
	End Function	
	Public Sub setIdOrder(strIdOrder)
		idOrder = strIdOrder
	End Sub
	
	Public Function getIdProd()
		getIdProd = idProd
	End Function	
	Public Sub setIdProd(strIdProd)
		idProd = strIdProd
	End Sub
	
	Public Function getIdDownProd()
		getIdDownProd = idDownProd
	End Function	
	Public Sub setIdDownProd(strIdDownProd)
		idDownProd = strIdDownProd
	End Sub
	
	Public Function getIdUser()
		getIdUser = idUser
	End Function	
	Public Sub setIdUser(strIdUser)
		idUser = strIdUser
	End Sub
	
	Public Function isActive()
		isActive = active
	End Function	
	Public Sub setIsActive(bolIsActive)
		active = bolIsActive
	End Sub
	
	Public Function getMaxNumDownload()
		getMaxNumDownload = maxNumDownload
	End Function	
	Public Sub setMaxNumDownload(strMaxNumDownload)
		maxNumDownload = strMaxNumDownload
	End Sub
	
	Public Function getInsertDate()
		getInsertDate = insertDate
	End Function	
	Public Sub setInsertDate(strInsertDate)
		insertDate = strInsertDate
	End Sub
	
	Public Function getExpireDate()
		getExpireDate = expireDate
	End Function	
	Public Sub setExpireDate(strExpireDate)
		expireDate = strExpireDate
	End Sub
	
	Public Function getDownloadCounter()
		getDownloadCounter = downloadCounter
	End Function	
	Public Sub setDownloadCounter(strDownloadCounter)
		downloadCounter = strDownloadCounter
	End Sub
	
	Public Function getDownloadDate()
		getDownloadDate = downloadDate
	End Function	
	Public Sub setDownloadDate(strDownloadDate)
		downloadDate = strDownloadDate
	End Sub
	
	Public Function isExpired()
		isExpired = true
		if(isNull(expireDate) OR expireDate="" OR expireDate = -1) then
			isExpired = false
		else
			if(DateDiff("n",expireDate,now()) < 0)then
				isExpired = false
			end if
		end if
	End Function	
	
	Public Function isMaxDownNum()
		isMaxDownNum = true		
		if(maxNumDownload = -1) then
			isMaxDownNum = false
		else		
			if(maxNumDownload > downloadCounter)then
				isMaxDownNum = false
			end if
		end if
	End Function		
	
	
	Public Function getFilePerOrdine(id_order)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, strID
		Set getFilePerOrdine = null		
		strSQL = "SELECT * FROM down_prod_x_order WHERE id_order=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objDownProd = new DownloadableProduct4OrderClass
				strID = objRS("id")		
				objDownProd.setID(strID)
				objDownProd.setIdOrder(objRS("id_order"))
				objDownProd.setIdProd(objRS("id_prod"))
				objDownProd.setIdDownProd(objRS("id_down_prod"))
				objDownProd.setIdUser(objRS("id_user"))
				objDownProd.setIsActive(objRS("active"))
				objDownProd.setMaxNumDownload(objRS("max_num_download"))
				objDownProd.setInsertDate(objRS("insert_date"))		
				objDownProd.setExpireDate(objRS("expire_date"))		
				objDownProd.setDownloadCounter(objRS("download_counter"))		
				objDownProd.setDownloadDate(objRS("download_date"))		
			
				objDict.add strID, objDownProd
				Set objFiles = nothing
				objRS.moveNext()
			loop
							
			Set getFilePerOrdine = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function	
	
	
	Public Function getFilePerProdotto(id_order, id_prodotto)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, strID
		Set getFilePerProdotto = null		
		strSQL = "SELECT * FROM down_prod_x_order WHERE id_order=? AND id_prod=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prodotto)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objDownProd = new DownloadableProduct4OrderClass
				strID = objRS("id")		
				objDownProd.setID(strID)
				objDownProd.setIdOrder(objRS("id_order"))
				objDownProd.setIdProd(objRS("id_prod"))
				objDownProd.setIdDownProd(objRS("id_down_prod"))
				objDownProd.setIdUser(objRS("id_user"))
				objDownProd.setIsActive(objRS("active"))
				objDownProd.setMaxNumDownload(objRS("max_num_download"))
				objDownProd.setInsertDate(objRS("insert_date"))		
				objDownProd.setExpireDate(objRS("expire_date"))		
				objDownProd.setDownloadCounter(objRS("download_counter"))		
				objDownProd.setDownloadDate(objRS("download_date"))		
			
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
	
	Public Function getFileByIDProdDown(id_order, id_prodotto, id_file)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDownProd
		getFileByIDProdDown = null		
		strSQL = "SELECT * FROM down_prod_x_order WHERE id_order=? AND id_prod=? AND id_down_prod=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_file)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objDownProd = new DownloadableProduct4OrderClass			
			objDownProd.setID(objRS("id"))
			objDownProd.setIdOrder(objRS("id_order"))
			objDownProd.setIdProd(objRS("id_prod"))
			objDownProd.setIdDownProd(objRS("id_down_prod"))
			objDownProd.setIdUser(objRS("id_user"))
			objDownProd.setIsActive(objRS("active"))
			objDownProd.setMaxNumDownload(objRS("max_num_download"))
			objDownProd.setInsertDate(objRS("insert_date"))		
			objDownProd.setExpireDate(objRS("expire_date"))		
			objDownProd.setDownloadCounter(objRS("download_counter"))		
			objDownProd.setDownloadDate(objRS("download_date"))
							
			Set getFileByIDProdDown = objDownProd
			Set objDownProd = nothing				
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
		Dim objDB, strSQL, objRS, objConn, objDownProd, strID
		getFileByID = null		
		strSQL = "SELECT * FROM down_prod_x_order WHERE id=?;"
				
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
			Set objDownProd = new DownloadableProduct4OrderClass
			strID = objRS("id")		
			objDownProd.setID(strID)
			objDownProd.setIdOrder(objRS("id_order"))
			objDownProd.setIdProd(objRS("id_prod"))
			objDownProd.setIdDownProd(objRS("id_down_prod"))
			objDownProd.setIdUser(objRS("id_user"))
			objDownProd.setIsActive(objRS("active"))
			objDownProd.setMaxNumDownload(objRS("max_num_download"))
			objDownProd.setInsertDate(objRS("insert_date"))		
			objDownProd.setExpireDate(objRS("expire_date"))		
			objDownProd.setDownloadCounter(objRS("download_counter"))		
			objDownProd.setDownloadDate(objRS("download_date"))
							
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

	'********************************************* METODI DAO *********************************************
	Public Function insertDownProd(IdOrder, IdProd, IdDownProd, IdUser, IsActive, MaxNumDownload, InsertDate, ExpireDate, DownloadCounter, DownloadDate, objConn)
		on error resume next
		insertDownProd = -1
		
		Dim strSQL, objRS, dtData_ins, dtExpireDate, dtDownloadDate
		
		InsertDate = convertDate(InsertDate)

		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			dtExpireDate = convertDate(ExpireDate)
		else	
			'controllo il tipo di database in uso
			if (Application("dbType") = 0) then
				dtExpireDate = "null"
			else		
				dtExpireDate = "'0000-00-00 00:00:00'"
			end if			
		end if

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
		
		strSQL = "INSERT INTO down_prod_x_order(id_order, id_prod, id_down_prod, id_user, active, max_num_download, insert_date, expire_date, download_counter, download_date) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,"
		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			strSQL = strSQL & "?"
		else
			strSQL = strSQL & dtExpireDate
		end if		
		strSQL = strSQL & ",?,"
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
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdDownProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,18,1,,IsActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,MaxNumDownload)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,InsertDate)
		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtExpireDate)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,DownloadCounter)
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtDownloadDate)
		end if
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(down_prod_x_order.id) as id FROM down_prod_x_order;")
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

	Public Function insertDownProdNoTransaction(IdOrder, IdProd, IdDownProd, IdUser, IsActive, MaxNumDownload, InsertDate, ExpireDate, DownloadCounter, DownloadDate)
		on error resume next
		insertDownProdNoTransaction = -1
		
		Dim objDB, strSQL, objRS, objConn, dtData_ins, dtExpireDate, dtDownloadDate
		
		InsertDate = convertDate(InsertDate)

		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			dtExpireDate = convertDate(ExpireDate)
		else	
			'controllo il tipo di database in uso
			if (Application("dbType") = 0) then
				dtExpireDate = "null"
			else		
				dtExpireDate = "'0000-00-00 00:00:00'"
			end if			
		end if

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
		
		strSQL = "INSERT INTO down_prod_x_order(id_order, id_prod, id_down_prod, id_user, active, max_num_download, insert_date, expire_date, download_counter, download_date) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,"
		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			strSQL = strSQL & "?"
		else
			strSQL = strSQL & dtExpireDate
		end if		
		strSQL = strSQL & ",?,"
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
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdDownProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,18,1,,IsActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,MaxNumDownload)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,InsertDate)
		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtExpireDate)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,DownloadCounter)
		if not(isNull(DownloadDate))  AND NOT (DownloadDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtDownloadDate)
		end if
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(down_prod_x_order.id) as id FROM down_prod_x_order;")
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
		
	Public Sub modifyDownProd(idThis, IdOrder, IdProd, IdDownProd, IdUser, IsActive, MaxNumDownload, InsertDate, ExpireDate, DownloadCounter, DownloadDate, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, dtExpireDate, dtDownloadDate
		
		InsertDate = convertDate(InsertDate)

		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			dtExpireDate = convertDate(ExpireDate)
		else	
			'controllo il tipo di database in uso
			if (Application("dbType") = 0) then
				dtExpireDate = "null"
			else		
				dtExpireDate = "'0000-00-00 00:00:00'"
			end if			
		end if

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
		
		strSQL = "UPDATE down_prod_x_order SET "
		strSQL = strSQL & "id_order=?,"
		strSQL = strSQL & "id_prod=?,"
		strSQL = strSQL & "id_down_prod=?,"
		strSQL = strSQL & "id_user=?,"
		strSQL = strSQL & "active=?,"
		strSQL = strSQL & "max_num_download=?,"
		strSQL = strSQL & "insert_date=?,"
		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			strSQL = strSQL & "expire_date=?,"
		else
			strSQL = strSQL & "expire_date="&dtExpireDate&","
		end if
		strSQL = strSQL & "download_counter=?,"
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
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdDownProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,18,1,,IsActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,MaxNumDownload)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,InsertDate)
		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtExpireDate)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,DownloadCounter)
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
		
	Public Sub modifyDownProdNoTransaction(idThis, IdOrder, IdProd, IdDownProd, IdUser, IsActive, MaxNumDownload, InsertDate, ExpireDate, DownloadCounter, DownloadDate)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, dtExpireDate, dtDownloadDate

		InsertDate = convertDate(InsertDate)

		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			dtExpireDate = convertDate(ExpireDate)
		else	
			'controllo il tipo di database in uso
			if (Application("dbType") = 0) then
				dtExpireDate = "null"
			else		
				dtExpireDate = "'0000-00-00 00:00:00'"
			end if			
		end if

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

		strSQL = "UPDATE down_prod_x_order SET "
		strSQL = strSQL & "id_order=?,"
		strSQL = strSQL & "id_prod=?,"
		strSQL = strSQL & "id_down_prod=?,"
		strSQL = strSQL & "id_user=?,"
		strSQL = strSQL & "active=?,"
		strSQL = strSQL & "max_num_download=?,"
		strSQL = strSQL & "insert_date=?,"
		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			strSQL = strSQL & "expire_date=?,"
		else
			strSQL = strSQL & "expire_date="&dtExpireDate&","
		end if
		strSQL = strSQL & "download_counter=?,"
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
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdDownProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,IdUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,18,1,,IsActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,MaxNumDownload)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,InsertDate)
		if not(isNull(ExpireDate))  AND NOT (ExpireDate = "")  then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtExpireDate)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,DownloadCounter)
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
		
	Public Sub deleteDownProd(idThis, objConn)
		on error resume next
		Dim objDB, strSQLDelProdotto, objRS
		strSQLDelProdotto = "DELETE FROM down_prod_x_order WHERE id=?;"

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
		strSQLDelProdotto = "DELETE FROM down_prod_x_order WHERE id=?;"

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