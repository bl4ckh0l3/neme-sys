<%
Class PaymentFieldClass
	Private id
	Private idPayment
	Private idModulo
	Private nameField
	Private valueField
	Private matchField
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub
	
	Public Function getPaymentID()
		getPaymentID = idPayment
	End Function
	
	Public Sub setPaymentID(strPaymentID)
		idPayment = strPaymentID
	End Sub
	
	Public Function getModuleID()
		getModuleID = idModulo
	End Function
	
	Public Sub setModuleID(strModuleID)
		idModulo = strModuleID
	End Sub
	
	Public Function getNameField()
		getNameField = nameField
	End Function
	
	Public Sub setNameField(strNameField)
		nameField = strNameField
	End Sub
	
	Public Function getValueField()
		getValueField = valueField
	End Function
	
	Public Sub setValueField(strValueField)
		valueField = strValueField
	End Sub
	
	Public Function getMatchField()
		getMatchField = matchField
	End Function
	
	Public Sub setMatchField(strMatchField)
		matchField = strMatchField
	End Sub

	
	Public Function getListaPaymentFieldByIdPaymentAndModule(id_payment, id_module)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaPaymentFieldByIdPaymentAndModule = null	
		
		strSQL = "SELECT * FROM payment_field WHERE id_payment=? AND id_modulo=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_payment)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_module)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objPaymentField
			do while not objRS.EOF				
				Set objPaymentField = new PaymentFieldClass
				strID = objRS("id")
				objPaymentField.setID(strID)
				objPaymentField.setPaymentID(objRS("id_payment"))
				objPaymentField.setModuleID(objRS("id_modulo"))
				objPaymentField.setNameField(objRS("name"))	
				objPaymentField.setValueField(objRS("value"))			
				objPaymentField.setMatchField(objRS("match_field"))		
				objDict.add strID, objPaymentField
				objRS.moveNext()
			loop
			Set objPaymentField = nothing							
			Set getListaPaymentFieldByIdPaymentAndModule = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function getListaPaymentFieldNotMatch(id_payment, id_module)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaPaymentFieldNotMatch = null	
		
		strSQL = "SELECT * FROM payment_field WHERE id_payment=? AND id_modulo=? AND (match_field ='' OR match_field IS NULL);"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_payment)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_module)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objPaymentField
			do while not objRS.EOF				
				Set objPaymentField = new PaymentFieldClass
				strID = objRS("id")
				objPaymentField.setID(strID)
				objPaymentField.setPaymentID(objRS("id_payment"))
				objPaymentField.setModuleID(objRS("id_modulo"))
				objPaymentField.setNameField(objRS("name"))	
				objPaymentField.setValueField(objRS("value"))			
				objPaymentField.setMatchField(objRS("match_field"))		
				objDict.add strID, objPaymentField
				objRS.moveNext()
			loop
			Set objPaymentField = nothing							
			Set getListaPaymentFieldNotMatch = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function getListaPaymentFieldDoMatch(id_payment, id_module)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaPaymentFieldDoMatch = null	
		
		strSQL = "SELECT * FROM payment_field WHERE id_payment=? AND id_modulo=? AND match_field <>'' AND NOT match_field IS NULL;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_payment)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_module)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objPaymentField
			do while not objRS.EOF				
				Set objPaymentField = new PaymentFieldClass
				strID = objRS("id")
				objPaymentField.setID(strID)
				objPaymentField.setPaymentID(objRS("id_payment"))
				objPaymentField.setModuleID(objRS("id_modulo"))
				objPaymentField.setNameField(objRS("name"))	
				objPaymentField.setValueField(objRS("value"))			
				objPaymentField.setMatchField(objRS("match_field"))		
				objDict.add strID, objPaymentField
				objRS.moveNext()
			loop
			Set objPaymentField = nothing							
			Set getListaPaymentFieldDoMatch = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function getListaCheckinNotMatch(id_payment, id_module)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaCheckinNotMatch = null	
		
		strSQL = "SELECT * FROM payment_field WHERE id_payment=? AND id_modulo=? AND (match_field ='' OR match_field IS NULL);"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_payment)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_module)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objPaymentField
			do while not objRS.EOF				
				Set objPaymentField = new PaymentFieldClass
				strID = objRS("id")
				strName = objRS("name")
				objPaymentField.setID(strID)
				objPaymentField.setPaymentID(objRS("id_payment"))
				objPaymentField.setModuleID(objRS("id_modulo"))
				objPaymentField.setNameField(strName)	
				objPaymentField.setValueField(objRS("value"))			
				objPaymentField.setMatchField(objRS("match_field"))		
				objDict.add strName, objPaymentField
				objRS.moveNext()
			loop
			Set objPaymentField = nothing							
			Set getListaCheckinNotMatch = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function getListaCheckinDoMatch(id_payment, id_module)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaCheckinDoMatch = null	
		
		strSQL = "SELECT * FROM payment_field WHERE id_payment=? AND id_modulo=? AND match_field <>'' AND NOT match_field IS NULL;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_payment)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_module)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objPaymentField
			do while not objRS.EOF				
				Set objPaymentField = new PaymentFieldClass
				strID = objRS("id")
				strName = objRS("name")
				objPaymentField.setID(strID)
				objPaymentField.setPaymentID(objRS("id_payment"))
				objPaymentField.setModuleID(objRS("id_modulo"))
				objPaymentField.setNameField(strName)	
				objPaymentField.setValueField(objRS("value"))			
				objPaymentField.setMatchField(objRS("match_field"))		
				objDict.add strName, objPaymentField
				objRS.moveNext()
			loop
			Set objPaymentField = nothing							
			Set getListaCheckinDoMatch = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findPaymentFieldById(id, id_payment)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentFieldById = null		
		strSQL = "SELECT * FROM payment_field WHERE id=? AND id_payment=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_payment)
		Set objRS = objCommand.Execute()			
		
		if not(objRS.EOF) then
			Dim objPaymentField
			Set objPaymentField = new PaymentFieldClass
			strID = objRS("id")
			objPaymentField.setID(strID)
			objPaymentField.setPaymentID(objRS("id_payment"))
			objPaymentField.setModuleID(objRS("id_modulo"))
			objPaymentField.setNameField(objRS("name"))	
			objPaymentField.setValueField(objRS("value"))			
			objPaymentField.setMatchField(objRS("match_field"))			
			Set findPaymentFieldById = objPaymentField
			Set objPaymentField = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function findPaymentFieldByName(id_payment, id_module, nameField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentFieldByName = null		
		strSQL = "SELECT * FROM payment_field WHERE id_payment=? AND id_modulo=? AND name=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_payment)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_module)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,nameField)
		Set objRS = objCommand.Execute()				
		
		if not(objRS.EOF) then
			Dim objPaymentField
			Set objPaymentField = new PaymentFieldClass
			strID = objRS("id")
			objPaymentField.setID(strID)
			objPaymentField.setPaymentID(objRS("id_payment"))
			objPaymentField.setModuleID(objRS("id_modulo"))
			objPaymentField.setNameField(objRS("name"))	
			objPaymentField.setValueField(objRS("value"))			
			objPaymentField.setMatchField(objRS("match_field"))			
			Set findPaymentFieldByName = objPaymentField
			Set objPaymentField = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function		
		
	Public Function findPaymentFieldByMatch(id_payment, id_module, matchField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentFieldByMatch = null		
		strSQL = "SELECT * FROM payment_field WHERE id_payment=? AND id_modulo=? AND name=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_payment)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_module)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,matchField)
		Set objRS = objCommand.Execute()				
		
		if not(objRS.EOF) then
			Dim objPaymentField
			Set objPaymentField = new PaymentFieldClass
			strID = objRS("id")
			objPaymentField.setID(strID)
			objPaymentField.setPaymentID(objRS("id_payment"))
			objPaymentField.setModuleID(objRS("id_modulo"))
			objPaymentField.setNameField(objRS("name"))	
			objPaymentField.setValueField(objRS("value"))			
			objPaymentField.setMatchField(objRS("match_field"))			
			Set findPaymentFieldByMatch = objPaymentField
			Set objPaymentField = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertPaymentField(idPayment, idModule, strName, strValue, strMatchField, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO payment_field(id_payment, id_modulo, name, value, match_field) VALUES("
		strSQL = strSQL & "?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idPayment)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idModule)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strMatchField)
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
		
	Public Sub modifyPaymentField(id, idPayment, idModule, strName, strValue, strMatchField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE payment_field SET "
		strSQL = strSQL & "id_payment=?,"
		strSQL = strSQL & "id_modulo=?,"
		strSQL = strSQL & "name=?,"
		strSQL = strSQL & "value=?,"		
		strSQL = strSQL & "match_field=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idPayment)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idModule)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strMatchField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deletePaymentField(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM payment_field WHERE id=?;"

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
		
	Public Sub deletePaymentFieldList(id_payment, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM payment_field WHERE id_payment=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_payment)
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
	
	Public Function getListaMatchFields()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		Set getListaMatchFields = Server.CreateObject("Scripting.Dictionary")
		
		strSQL = "SELECT keyword, value FROM payment_fixed_app_field WHERE used=1;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim keyword, value
			do while not objRS.EOF
				keyword = objRS("keyword")
				value = objRS("value")
				objDict.add keyword, value
				objRS.moveNext()
			loop					
			Set getListaMatchFields = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function	
End Class
%>