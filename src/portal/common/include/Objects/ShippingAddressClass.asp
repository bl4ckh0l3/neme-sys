<%
Class ShippingAddressClass
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
	private companyClient

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

	Public Function isCompanyClient()
		isCompanyClient = companyClient
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
	
	Public Sub setCompanyClient(bolCompanyClient)
		companyClient = bolCompanyClient
	End Sub
	
	
	Public Function insertShippingAddress(idUser, strAddress, strName, strSurname, strCfiscVat, strCity, strZipCode, strCountry, strStateRegion, isCompanyClient, objConn)
		on error resume next
		insertShippingAddress = -1
		
		Dim objDB, strSQL, objRS	

		strSQL = "INSERT INTO shipping_address(id_user, address, name, surname, cfiscvat, city, zipCode, country, state_region, is_company_client) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"			
		end if
		strSQL = strSQL & "?);"

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
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,isCompanyClient)
		objCommand.Execute()
		Set objRS = objConn.Execute("SELECT max(shipping_address.id) as id FROM shipping_address")
		if not (objRS.EOF) then
			insertShippingAddress = objRS("id")	
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

	Public Sub modifyShippingAddress(id, idUser, strAddress, strName, strSurname, strCfiscVat, strCity, strZipCode, strCountry, strStateRegion, isCompanyClient, objConn)
		on error resume next
		Dim objDB, strSQL
		
		strSQL = "UPDATE shipping_address SET "
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
		strSQL = strSQL & ",is_company_client=?"
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
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,isCompanyClient)
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
	
	Public Function insertShippingAddressNoTransaction(idUser, strAddress, strName, strSurname, strCfiscVat, strCity, strZipCode, strCountry, strStateRegion, isCompanyClient)
		on error resume next
		insertShippingAddressNoTransaction = -1
		
		Dim objDB, strSQL, objRS, objConn		

		strSQL = "INSERT INTO shipping_address(id_user, address, name, surname, cfiscvat, city, zipCode, country, state_region, is_company_client) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"			
		end if
		strSQL = strSQL & "?);"
		
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
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,isCompanyClient)
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(shipping_address.id) as id FROM shipping_address")
		if not (objRS.EOF) then
			insertShippingAddressNoTransaction = objRS("id")	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
			
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function

	Public Sub modifyShippingAddressNoTransaction(id, idUser, strAddress, strName, strSurname, strCfiscVat, strCity, strZipCode, strCountry, strStateRegion, isCompanyClient)
		on error resume next
		Dim objDB, strSQL, objConn
		
		strSQL = "UPDATE shipping_address SET "
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
		strSQL = strSQL & ",is_company_client=?"
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
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,isCompanyClient)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
		
	Public Function deleteShippingAddress(id)
		on error resume next		
		deleteShippingAddress = true
		
		Dim objDB, strSQL, objRS, objRS2, objConn, strSQL2, strSQL3
		strSQL = "DELETE FROM shipping_address WHERE id=?;" 
		strSQL2 = "SELECT order_shipping_address.id_shipping FROM order_shipping_address WHERE order_shipping_address.id_shipping=?;"
		strSQL3 = "DELETE FROM order_shipping_address WHERE id_shipping=?;" 

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
			deleteShippingAddress = false				
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
	
	Public Function findShippingAddressByID(id)
		on error resume next
		
		findShippingAddressByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM shipping_address WHERE id=?;"

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
			Dim objShippingAddress
			Set objShippingAddress = new ShippingAddressClass
			objShippingAddress.setID(objRS("id"))
			objShippingAddress.setUserID(objRS("id_user"))
			objShippingAddress.setAddress(objRS("address"))
			objShippingAddress.setName(objRS("name"))
			objShippingAddress.setSurname(objRS("surname"))
			objShippingAddress.setCfiscVat(objRS("cfiscvat"))
			objShippingAddress.setCity(objRS("city"))
			objShippingAddress.setZipCode(objRS("zipCode"))
			objShippingAddress.setCountry(objRS("country"))	
			objShippingAddress.setStateRegion(objRS("state_region"))
			objShippingAddress.setCompanyClient(objRS("is_company_client"))		

			Set findShippingAddressByID = objShippingAddress
			Set objShippingAddress = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function
		
	Public Function findShippingAddressByUserID(strUser)
		on error resume next
		
		findShippingAddressByUserID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM shipping_address WHERE id_user=?;"

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
			Dim objShippingAddress
			Set objShippingAddress = new ShippingAddressClass
			objShippingAddress.setID(objRS("id"))
			objShippingAddress.setUserID(objRS("id_user"))
			objShippingAddress.setAddress(objRS("address"))
			objShippingAddress.setName(objRS("name"))
			objShippingAddress.setSurname(objRS("surname"))
			objShippingAddress.setCfiscVat(objRS("cfiscvat"))
			objShippingAddress.setCity(objRS("city"))
			objShippingAddress.setZipCode(objRS("zipCode"))
			objShippingAddress.setCountry(objRS("country"))
			objShippingAddress.setStateRegion(objRS("state_region"))
			objShippingAddress.setCompanyClient(objRS("is_company_client"))		
			
			Set findShippingAddressByUserID = objShippingAddress
			Set objShippingAddress = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertOrderShippingAddress(id_order, id_shipping, strAddress, strCity, strZipCode, strCountry, strStateRegion, isCompanyClient, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "INSERT INTO order_shipping_address(id_order, id_shipping, address, city, zipCode, country, state_region, is_company_client) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"			
		end if
		strSQL = strSQL & "?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_shipping)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,isCompanyClient)
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
	
	Public Sub insertOrderShippingAddressNoTransaction(id_order, id_shipping, strAddress, strCity, strZipCode, strCountry, strStateRegion, isCompanyClient)
		on error resume next
		Dim objDB, strSQL, objConn
		
		strSQL = "INSERT INTO order_shipping_address(id_order, id_shipping, address, city, zipCode, country, state_region, is_company_client) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"			
		end if
		strSQL = strSQL & "?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_shipping)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,isCompanyClient)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub modifyOrderShippingAddress(id_order, id_shipping, strAddress, strCity, strZipCode, strCountry, strStateRegion, isCompanyClient, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "UPDATE order_shipping_address SET "
		strSQL = strSQL & "id_order=?,"
		strSQL = strSQL & "id_shipping=?,"
		strSQL = strSQL & "address=?,"
		strSQL = strSQL & "city=?,"
		strSQL = strSQL & "zipCode=?,"
		strSQL = strSQL & "country=?"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & ",state_region=NULL"
		else
			strSQL = strSQL & ",state_region=?"			
		end if
		strSQL = strSQL & ",is_company_client=?"
		strSQL = strSQL & " WHERE id_order=?;" 	

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_shipping)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,isCompanyClient)
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
	
	Public Sub modifyOrderShippingAddressNoTransaction(id_order, id_shipping, strAddress, strCity, strZipCode, strCountry, strStateRegion, isCompanyClient)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE order_shipping_address SET "
		strSQL = strSQL & "id_order=?,"
		strSQL = strSQL & "id_shipping=?,"
		strSQL = strSQL & "address=?,"
		strSQL = strSQL & "city=?,"
		strSQL = strSQL & "zipCode=?,"
		strSQL = strSQL & "country=?"
		if(isNull(strStateRegion) OR strStateRegion = "") then
			strSQL = strSQL & ",state_region=NULL"
		else
			strSQL = strSQL & ",state_region=?"			
		end if
		strSQL = strSQL & ",is_company_client=?"
		strSQL = strSQL & " WHERE id_order=?;" 	
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_shipping)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strAddress)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCity)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,20,strZipCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCountry)
		if not(isNull(strStateRegion)) AND not(strStateRegion = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStateRegion)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,isCompanyClient)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_order)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteOrderShippingAddress(id_shipping, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM order_shipping_address WHERE id_shipping=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_shipping)
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
		
	Public Sub deleteOrderShippingAddressNoTransaction(id_shipping)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM order_shipping_address WHERE id_shipping=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_shipping)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Function getOrderShippingAddress(id_order)		
		on error resume next
		
		getOrderShippingAddress = null
		
		Dim objDB, strSQL, objRS, objConn	
		strSQL = "SELECT order_shipping_address.*, shipping_address.id_user, shipping_address.name, shipping_address.surname, shipping_address.cfiscvat FROM shipping_address INNER JOIN order_shipping_address ON shipping_address.id = order_shipping_address.id_shipping WHERE order_shipping_address.id_order=?;"

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
			Dim objShippingAddress
			Set objShippingAddress = new ShippingAddressClass
			objShippingAddress.setID(objRS("id_shipping"))
			objShippingAddress.setUserID(objRS("id_user"))
			objShippingAddress.setOrderID(objRS("id_order"))
			objShippingAddress.setAddress(objRS("address"))
			objShippingAddress.setName(objRS("name"))
			objShippingAddress.setSurname(objRS("surname"))
			objShippingAddress.setCfiscVat(objRS("cfiscvat"))
			objShippingAddress.setCity(objRS("city"))
			objShippingAddress.setZipCode(objRS("zipCode"))
			objShippingAddress.setCountry(objRS("country"))
			objShippingAddress.setStateRegion(objRS("state_region"))	
			objShippingAddress.setCompanyClient(objRS("is_company_client"))	
			
			Set getOrderShippingAddress = objShippingAddress
			Set objShippingAddress = Nothing
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