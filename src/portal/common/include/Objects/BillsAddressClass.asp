<%
Class BillsAddressClass
	Private id
	Private idOrder
	Private idUser
	private name
	private surname
	private cfiscvat
	private address
	private city
	private zipCode
	private country
	private stateRegion

	'****************** GET FUNCTIONS ***************
	
	Public Function getID()
		getID = id
	End Function
	
	Public Function getOrderID()
		getOrderID = idOrder
	End Function
	
	Public Function getUserID()
		getUserID = idUser
	End Function
	
	Public Function getName()
		getName = name
	End Function
	
	Public Function getSurname()
		getSurname = surname
	End Function
	
	Public Function getCfiscVat()
		getCfiscVat = cfiscvat
	End Function
	
	Public Function getAddress()
		getAddress = address
	End Function
	
	Public Function getCity()
		getCity = city
	End Function
	
	Public Function getZipCode()
		getZipCode = zipCode
	End Function
	
	Public Function getCountry()
		getCountry = country
	End Function
	
	Public Function getStateRegion()
		getStateRegion = stateRegion
	End Function
	
	
	'****************** SET FUNCTIONS ***************
				
	Public Sub setID(strID)
		id = strID
	End Sub
				
	Public Sub setOrderID(strOrderID)
		idOrder = strOrderID
	End Sub
	
	Public Sub setUserID(strUserID)
		idUser = strUserID
	End Sub
	
	Public Sub setAddress(strAddress)
		address = strAddress
	End Sub
	
	Public Sub setName(strName)
		name = strName
	End Sub
	
	Public Sub setSurname(strSurname)
		surname = strSurname
	End Sub
	
	Public Sub setCfiscVat(strCfiscVat)
		cfiscvat = strCfiscVat
	End Sub
	
	Public Sub setCity(strCity)
		city = strCity
	End Sub
	
	Public Sub setZipCode(strZipCode)
		zipCode = strZipCode
	End Sub
	
	Public Sub setCountry(strCountry)
		country = strCountry
	End Sub
	
	Public Sub setStateRegion(strStateRegion)
		stateRegion = strStateRegion
	End Sub
	
	
	Public Function insertBillsAddress(idUser, strAddress, strName, strSurname, strCfiscVat, strCity, strZipCode, strCountry, strStateRegion, objConn)
		on error resume next
		insertBillsAddress = -1
		
		Dim objDB, strSQL, objRS	

		strSQL = "INSERT INTO bills_address(id_user, address, name, surname, cfiscvat, city, zipCode, country, state_region) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & "NULL);"
		else
			strSQL = strSQL & "?);"			
		end if

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strSurname)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,16,strCfiscVat)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Execute()
		Set objRS = objConn.Execute("SELECT max(bills_address.id) as id FROM bills_address")
		if not (objRS.EOF) then
			insertBillsAddress = objRS("id")	
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

	Public Sub modifyBillsAddress(id, idUser, strAddress, strName, strSurname, strCfiscVat, strCity, strZipCode, strCountry, strStateRegion, objConn)
		on error resume next
		Dim objDB, strSQL
		
		strSQL = "UPDATE bills_address SET "
		strSQL = strSQL & "id_user=?,"
		strSQL = strSQL & "address=?,"
		strSQL = strSQL & "name=?,"
		strSQL = strSQL & "surname=?,"
		strSQL = strSQL & "cfiscvat=?,"
		strSQL = strSQL & "city=?,"
		strSQL = strSQL & "zipCode=?,"
		strSQL = strSQL & "country=?"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & ",state_region=NULL"
		else
			strSQL = strSQL & ",state_region=?"			
		end if
		strSQL = strSQL & " WHERE id=?;" 

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strSurname)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,16,strCfiscVat)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
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
	
	Public Function insertBillsAddressNoTransaction(idUser, strAddress, strName, strSurname, strCfiscVat, strCity, strZipCode, strCountry, strStateRegion)
		on error resume next
		insertBillsAddressNoTransaction = -1
		
		Dim objDB, strSQL, objRS, objConn		

		strSQL = "INSERT INTO bills_address(id_user, address, name, surname, cfiscvat, city, zipCode, country, state_region) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & "NULL);"
		else
			strSQL = strSQL & "?);"			
		end if
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strSurname)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,16,strCfiscVat)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Execute()
		Set objRS = objConn.Execute("SELECT max(bills_address.id) as id FROM bills_address")
		if not (objRS.EOF) then
			insertBillsAddressNoTransaction = objRS("id")	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
			
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function

	Public Sub modifyBillsAddressNoTransaction(id, idUser, strAddress, strName, strSurname, strCfiscVat, strCity, strZipCode, strCountry, strStateRegion)
		on error resume next
		Dim objDB, strSQL, objConn
		
		strSQL = "UPDATE bills_address SET "
		strSQL = strSQL & "id_user=?,"
		strSQL = strSQL & "address=?,"
		strSQL = strSQL & "name=?,"
		strSQL = strSQL & "surname=?,"
		strSQL = strSQL & "cfiscvat=?,"
		strSQL = strSQL & "city=?,"
		strSQL = strSQL & "zipCode=?,"
		strSQL = strSQL & "country=?"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & ",state_region=NULL"
		else
			strSQL = strSQL & ",state_region=?"			
		end if
		strSQL = strSQL & " WHERE id=?;" 
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()			

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strSurname)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,16,strCfiscVat)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
		
	Public Function deleteBillsAddress(id)
		on error resume next		
		deleteBillsAddress = true
		
		Dim objDB, strSQL, objRS, objRS2, objConn, strSQL2, strSQL3
		strSQL = "DELETE FROM bills_address WHERE id=?;" 
		strSQL2 = "SELECT order_bills_address.id_bills FROM order_bills_address WHERE order_bills_address.id_bills=?;"
		strSQL3 = "DELETE FROM order_bills_address WHERE id_bills=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		
		Dim objCommand, objCommand2, objCommand3
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		Set objCommand3 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand3.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand3.CommandType=1
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQL2
		objCommand3.CommandText = strSQL3
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)

		objConn.BeginTrans
		
		Set objRS = objCommand2.Execute()
		if not(objRS.EOF) then							
			deleteBillsAddress = false				
		else
			if(Application("use_innodb_table") = 0) then
				objCommand3.Execute()
			end if
			objCommand.Execute()
		end if
		
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
	
	Public Function findBillsAddressByID(id)
		on error resume next
		
		findBillsAddressByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM bills_address WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then
			Dim objBillsAddress
			Set objBillsAddress = new BillsAddressClass
			objBillsAddress.setID(objRS("id"))
			objBillsAddress.setUserID(objRS("id_user"))
			objBillsAddress.setAddress(objRS("address"))
			objBillsAddress.setName(objRS("name"))
			objBillsAddress.setSurname(objRS("surname"))
			objBillsAddress.setCfiscVat(objRS("cfiscvat"))
			objBillsAddress.setCity(objRS("city"))
			objBillsAddress.setZipCode(objRS("zipCode"))
			objBillsAddress.setCountry(objRS("country"))	
			objBillsAddress.setStateRegion(objRS("state_region"))

			Set findBillsAddressByID = objBillsAddress
			Set objBillsAddress = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function
		
	Public Function findBillsAddressByUserID(strUser)
		on error resume next
		
		findBillsAddressByUserID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM bills_address WHERE id_user=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strUser)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then
			Dim objBillsAddress
			Set objBillsAddress = new BillsAddressClass
			objBillsAddress.setID(objRS("id"))
			objBillsAddress.setUserID(objRS("id_user"))
			objBillsAddress.setAddress(objRS("address"))
			objBillsAddress.setName(objRS("name"))
			objBillsAddress.setSurname(objRS("surname"))
			objBillsAddress.setCfiscVat(objRS("cfiscvat"))
			objBillsAddress.setCity(objRS("city"))
			objBillsAddress.setZipCode(objRS("zipCode"))
			objBillsAddress.setCountry(objRS("country"))
			objBillsAddress.setStateRegion(objRS("state_region"))	
			
			Set findBillsAddressByUserID = objBillsAddress
			Set objBillsAddress = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertOrderBillsAddress(id_order, id_bills, strAddress, strCity, strZipCode, strCountry, strStateRegion, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "INSERT INTO order_bills_address(id_order, id_bills, address, city, zipCode, country, state_region) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & "NULL);"
		else
			strSQL = strSQL & "?);"			
		end if

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_bills)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
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
	
	Public Sub insertOrderBillsAddressNoTransaction(id_order, id_bills, strAddress, strCity, strZipCode, strCountry, strStateRegion)
		on error resume next
		Dim objDB, strSQL, objConn
		
		strSQL = "INSERT INTO order_bills_address(id_order, id_bills, address, city, zipCode, country, state_region) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & "NULL);"
		else
			strSQL = strSQL & "?);"			
		end if
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_bills)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub modifyOrderBillsAddress(id_order, id_bills, strAddress, strCity, strZipCode, strCountry, strStateRegion, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "UPDATE order_bills_address SET "
		strSQL = strSQL & "id_order=?,"
		strSQL = strSQL & "id_bills=?,"
		strSQL = strSQL & "address=?,"
		strSQL = strSQL & "city=?,"
		strSQL = strSQL & "zipCode=?,"
		strSQL = strSQL & "country=?"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & ",state_region=NULL"
		else
			strSQL = strSQL & ",state_region=?"			
		end if
		strSQL = strSQL & " WHERE id_order=?;" 	

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_bills)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
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
	
	Public Sub modifyOrderBillsAddressNoTransaction(id_order, id_bills, strAddress, strCity, strZipCode, strCountry, strStateRegion)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE order_bills_address SET "
		strSQL = strSQL & "id_order=?,"
		strSQL = strSQL & "id_bills=?,"
		strSQL = strSQL & "address=?,"
		strSQL = strSQL & "city=?,"
		strSQL = strSQL & "zipCode=?,"
		strSQL = strSQL & "country=?"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & ",state_region=NULL"
		else
			strSQL = strSQL & ",state_region=?"			
		end if
		strSQL = strSQL & " WHERE id_order=?;" 	
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_bills)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteOrderBillsAddress(id_bills, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM order_bills_address WHERE id_bills=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_bills)
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
		
	Public Sub deleteOrderBillsAddressNoTransaction(id_bills)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM order_bills_address WHERE id_bills=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_bills)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Function getOrderBillsAddress(id_order)
		
		on error resume next
		
		getOrderBillsAddress = null
		
		Dim objDB, strSQL, objRS, objConn	
		strSQL = "SELECT order_bills_address.*, bills_address.id_user, bills_address.name, bills_address.surname, bills_address.cfiscvat FROM bills_address INNER JOIN order_bills_address ON bills_address.id = order_bills_address.id_bills WHERE order_bills_address.id_order=?;"

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
			Dim objBillsAddress
			Set objBillsAddress = new BillsAddressClass
			objBillsAddress.setID(objRS("id_bills"))
			objBillsAddress.setUserID(objRS("id_user"))
			objBillsAddress.setOrderID(objRS("id_order"))
			objBillsAddress.setAddress(objRS("address"))
			objBillsAddress.setName(objRS("name"))
			objBillsAddress.setSurname(objRS("surname"))
			objBillsAddress.setCfiscVat(objRS("cfiscvat"))
			objBillsAddress.setCity(objRS("city"))
			objBillsAddress.setZipCode(objRS("zipCode"))
			objBillsAddress.setCountry(objRS("country"))	
			objBillsAddress.setStateRegion(objRS("state_region"))
			
			Set getOrderBillsAddress = objBillsAddress
			Set objBillsAddress = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function		
	

				
	public Sub toString()
		response.write (id & ", " & idOrder & ", " & idUser & ", " & address & ", " & city & ", " & zipCode & ", " & country)
	end Sub
End Class
%>