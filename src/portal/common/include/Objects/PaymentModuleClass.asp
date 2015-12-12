<%
Class PaymentModuleClass
	Private id
	Private nameModulo
	Private directory
	Private logo
	Private insertPage
	Private checkoutPage
	Private checkinPage
	Private checkinFaultPage
	Private idOrdineField
	Private ipProvider
	
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub
	
	Public Function getNameModulo()
		getNameModulo = nameModulo
	End Function
	
	Public Sub setNameModulo(strNameModulo)
		nameModulo = strNameModulo
	End Sub
	
	Public Function getDirectory()
		getDirectory = directory
	End Function
	
	Public Sub setDirectory(strDirectory)
		directory = strDirectory
	End Sub
	
	Public Function getLogo()
		getLogo = logo
	End Function
	
	Public Sub setLogo(strLogo)
		logo = strLogo
	End Sub
	
	Public Function getInsertPage()
		getInsertPage = insertPage
	End Function
	
	Public Sub setInsertPage(strInsertPage)
		insertPage = strInsertPage
	End Sub
	
	Public Function getCheckoutPage()
		getCheckoutPage = checkoutPage
	End Function
	
	Public Sub setCheckoutPage(strCheckoutPage)
		checkoutPage = strCheckoutPage
	End Sub
	
	Public Function getCheckinPage()
		getCheckinPage = checkinPage
	End Function
	
	Public Sub setCheckinPage(strCheckinPage)
		checkinPage = strCheckinPage
	End Sub
	
	Public Function getCheckinFaultPage()
		getCheckinFaultPage = checkinFaultPage
	End Function
	
	Public Sub setCheckinFaultPage(strCheckinFaultPage)
		checkinFaultPage = strCheckinFaultPage
	End Sub
	
	Public Function getIdOrdineField()
		getIdOrdineField = idOrdineField
	End Function
	
	Public Sub setIdOrdineField(ordineField)
		idOrdineField = ordineField
	End Sub
	
	Public Function getIpProvider()
		getIpProvider = ipProvider
	End Function
	
	Public Sub setIpProvider(numIpProvider)
		ipProvider = numIpProvider
	End Sub
	
		
	Public Function getListaPaymentModuli()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaPaymentModuli = null		
		strSQL = "SELECT * FROM payment_modulo;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objPaymentModule
			do while not objRS.EOF				
				Set objPaymentModule = new PaymentModuleClass
				strID = objRS("id")
				objPaymentModule.setID(strID)
				objPaymentModule.setNameModulo(objRS("name"))	
				objPaymentModule.setDirectory(objRS("directory"))	
				objPaymentModule.setLogo(objRS("logo"))	
				objPaymentModule.setInsertPage(objRS("insert_page"))		
				objPaymentModule.setCheckoutPage(objRS("checkout_page"))	
				objPaymentModule.setCheckinPage(objRS("checkin_page"))			
				objPaymentModule.setCheckinFaultPage(objRS("checkin_fault_page"))	
				objPaymentModule.setIdOrdineField(objRS("id_ordine_field"))	
				objPaymentModule.setIpProvider(objRS("ip_provider"))	
				objDict.add strID, objPaymentModule
				objRS.moveNext()
			loop
			Set objPaymentModule = nothing							
			Set getListaPaymentModuli = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findPaymentModuloByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentModuloByID = null		
		strSQL = "SELECT * FROM payment_modulo WHERE id=?;"
		
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
			Dim objPaymentModule
			Set objPaymentModule = new PaymentModuleClass
			objPaymentModule.setID(objRS("id"))
			objPaymentModule.setNameModulo(objRS("name"))	
			objPaymentModule.setDirectory(objRS("directory"))	
			objPaymentModule.setLogo(objRS("logo"))	
			objPaymentModule.setInsertPage(objRS("insert_page"))		
			objPaymentModule.setCheckoutPage(objRS("checkout_page"))	
			objPaymentModule.setCheckinPage(objRS("checkin_page"))	
			objPaymentModule.setCheckinFaultPage(objRS("checkin_fault_page"))		
			objPaymentModule.setIdOrdineField(objRS("id_ordine_field"))	
			objPaymentModule.setIpProvider(objRS("ip_provider"))	
			Set findPaymentModuloByID = objPaymentModule
			Set objPaymentModule = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findPaymentByDesc(nameModule)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentByDesc = null		
		strSQL = "SELECT * FROM payment_modulo WHERE name =?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,45,nameModule)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objPaymentModule
			Set objPaymentModule = new PaymentModuleClass
			objPaymentModule.setID(objRS("id"))
			objPaymentModule.setNameModulo(objRS("name"))	
			objPaymentModule.setDirectory(objRS("directory"))	
			objPaymentModule.setLogo(objRS("logo"))	
			objPaymentModule.setInsertPage(objRS("insert_page"))		
			objPaymentModule.setCheckoutPage(objRS("checkout_page"))	
			objPaymentModule.setCheckinPage(objRS("checkin_page"))	
			objPaymentModule.setCheckinFaultPage(objRS("checkin_fault_page"))			
			objPaymentModule.setIdOrdineField(objRS("id_ordine_field"))	
			objPaymentModule.setIpProvider(objRS("ip_provider"))	
			Set findPaymentByDesc = objPaymentModule
			Set objPaymentModule = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertPaymentModule(strNameModulo, strDirectory, strLogo, strInsertPage, strCheckoutPage, strCheckinPage, strCheckinFaultPage, idOrdineField, ipProvider)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO payment_modulo(name, directory, logo, insert_page, checkout_page, checkin_page, checkin_fault_page, id_ordine_field, ip_provider) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,45,strNameModulo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strDirectory)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strLogo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strInsertPage)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCheckoutPage)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCheckinPage)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCheckinFaultPage)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idOrdineField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,ipProvider)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyPaymentModule(id, strNameModulo, strDirectory, strLogo, strInsertPage, strCheckoutPage, strCheckinPage, strCheckinFaultPage, idOrdineField, ipProvider)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE payment_modulo SET "
		strSQL = strSQL & "name=?,"
		strSQL = strSQL & "directory=?,"
		strSQL = strSQL & "logo=?,"		
		strSQL = strSQL & "insert_page=?,"		
		strSQL = strSQL & "checkout_page=?,"		
		strSQL = strSQL & "checkin_page=?,"			
		strSQL = strSQL & "checkin_fault_page=?,"		
		strSQL = strSQL & "id_ordine_field=?,"	
		strSQL = strSQL & "ip_provider=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,45,strNameModulo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strDirectory)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strLogo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strInsertPage)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCheckoutPage)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCheckinPage)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strCheckinFaultPage)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idOrdineField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,ipProvider)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deletePaymentModule(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM payment_modulo WHERE id=?;"

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
	
End Class
%>