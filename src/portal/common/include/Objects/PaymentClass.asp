<%
Class PaymentClass
	Private id
	Private descrizione
	Private keyword_multilingua
	Private dati_pagamento
	Private commission
	Private commission_type
	Private url
	Private id_payment_module
	Private active
	Private payment_type
	
	
	Public Function getPaymentID()
		getPaymentID = id
	End Function
	
	Public Sub setPaymentID(strID)
		id = strID
	End Sub
	
	Public Function getKeywordMultilingua()
		getKeywordMultilingua = keyword_multilingua
	End Function
	
	Public Sub setKeywordMultilingua(strKeywordMultilingua)
		keyword_multilingua = strKeywordMultilingua
	End Sub
	
	Public Function getDescrizione()
		getDescrizione = descrizione
	End Function
	
	Public Sub setDescrizione(strDesc)
		descrizione = strDesc
	End Sub
	
	Public Function getDatiPagamento()
		getDatiPagamento = dati_pagamento
	End Function
	
	Public Sub setDatiPagamento(strDatiPagamento)
		dati_pagamento = strDatiPagamento
	End Sub
	
	Public Function getCommission()
		getCommission = commission
	End Function
	
	Public Sub setCommission(dblCommission)
		commission = dblCommission
	End Sub
	
	Public Function getCommissionType()
		getCommissionType = commission_type
	End Function
	
	Public Sub setCommissionType(intCommissionType)
		commission_type = intCommissionType
	End Sub
	
	Public Function getURL()
		getURL = url
	End Function
	
	Public Sub setURL(strURL)
		url = strURL
	End Sub
	
	Public Function getAttivo()
		getAttivo = active
	End Function
	
	Public Sub setAttivo(strAttivo)
		active = strAttivo
	End Sub	
	
	Public Function getPaymentModuleID()
		getPaymentModuleID = id_payment_module
	End Function
	
	Public Sub setPaymentModuleID(strModuleID)
		id_payment_module = strModuleID
	End Sub
	
	Public Function getPaymentType()
		getPaymentType = payment_type
	End Function
	
	Public Sub setPaymentType(bolPaymentType)
		payment_type = bolPaymentType
	End Sub	

	
	Public Function getImportoCommissione(dblAmount)
		commission = CDbl(commission)
		if(commission_type = 2) then
			importo = CDbl(dblAmount) * (commission / 100)
		else
			importo = commission
		end if
		
		getImportoCommissione = importo
	End Function
		
		
	Public Function getListaPayment(active, paymentType)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaPayment = null		
		
		strSQL = "SELECT * FROM payment_type"

		if not(isNull(active)) OR not(isNull(paymentType)) then
			strSQL = strSQL & " WHERE"
			
			if not(isNull(active)) then
			strSQL = strSQL & " AND activate=?"
			end if
			
			if not(isNull(paymentType)) then
			strSQL = strSQL & " AND payment_type=?"
			end if			
		end if
		
		strSQL = strSQL & " ORDER BY descrizione"
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"


		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if not(isNull(active)) then
			objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,active)
		end if
		if not(isNull(paymentType)) then
			objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,paymentType)
		end if
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objPayment
			do while not objRS.EOF				
				Set objPayment = new PaymentClass
				strID = objRS("id")
				objPayment.setPaymentID(strID)
				objPayment.setKeywordMultilingua(objRS("keyword_multilingua"))	
				objPayment.setDescrizione(objRS("descrizione"))	
				objPayment.setDatiPagamento(objRS("dati_pagamento"))	
				objPayment.setCommission(objRS("commission"))	
				objPayment.setCommissionType(objRS("commission_type"))
				objPayment.setURL(objRS("url"))	
				objPayment.setPaymentModuleID(objRS("id_modulo"))	
				objPayment.setAttivo(objRS("activate"))	
				objPayment.setPaymentType(objRS("payment_type"))	
				objDict.add strID, objPayment
				objRS.moveNext()
			loop
			Set objPayment = nothing							
			Set getListaPayment = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findPaymentByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentByID = null		
		strSQL = "SELECT * FROM payment_type WHERE id=?;"
		
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
			Dim objPayment
			Set objPayment = new PaymentClass
			objPayment.setPaymentID(objRS("id"))
			objPayment.setKeywordMultilingua(objRS("keyword_multilingua"))	
			objPayment.setDescrizione(objRS("descrizione"))	
			objPayment.setDatiPagamento(objRS("dati_pagamento"))		
			objPayment.setCommission(objRS("commission"))	
			objPayment.setCommissionType(objRS("commission_type"))	
			objPayment.setURL(objRS("url"))	
			objPayment.setPaymentModuleID(objRS("id_modulo"))				
			objPayment.setAttivo(objRS("activate"))	
			objPayment.setPaymentType(objRS("payment_type"))		
			Set findPaymentByID = objPayment
			Set objPayment = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findPaymentByDesc(description)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentByDesc = null		
		strSQL = "SELECT * FROM payment_type WHERE descrizione LIKE ?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,"%"&description&"%")
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objPayment
			Set objPayment = new PaymentClass
			objPayment.setPaymentID(objRS("id"))			
			objPayment.setKeywordMultilingua(objRS("keyword_multilingua"))	
			objPayment.setDescrizione(objRS("descrizione"))	
			objPayment.setDatiPagamento(objRS("dati_pagamento"))	
			objPayment.setCommission(objRS("commission"))	
			objPayment.setCommissionType(objRS("commission_type"))							
			objPayment.setURL(objRS("url"))					
			objPayment.setPaymentModuleID(objRS("id_modulo"))	
			objPayment.setAttivo(objRS("activate"))	
			objPayment.setPaymentType(objRS("payment_type"))			
			Set findPaymentByDesc = objPayment
			Set objPayment = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertPayment(strKeywordMultilingua, strDescrizione, dati_pagamento, commission, commission_type, url, id_modulo, activate, payment_type)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO payment_type(keyword_multilingua, descrizione, dati_pagamento,  commission, commission_type, url, id_modulo, activate, payment_type) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,?);"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strKeywordMultilingua)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,dati_pagamento)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(commission))
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,commission_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,url)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_modulo)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activate)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,payment_type)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyPayment(id, strKeywordMultilingua, strDescrizione, dati_pagamento, commission, commission_type, url, id_modulo, activate, payment_type, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE payment_type SET "
		strSQL = strSQL & "keyword_multilingua=?,"
		strSQL = strSQL & "descrizione=?,"
		strSQL = strSQL & "dati_pagamento=?,"	
		strSQL = strSQL & "commission=?,"	
		strSQL = strSQL & "commission_type=?,"		
		strSQL = strSQL & "url=?,"				
		strSQL = strSQL & "id_modulo=?,"	
		strSQL = strSQL & "activate=?,"
		strSQL = strSQL & "payment_type=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strKeywordMultilingua)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,dati_pagamento)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(commission))
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,commission_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,url)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_modulo)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activate)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,payment_type)
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
		
	Public Sub deletePayment(id)
		on error resume next
		Dim objDB, strSQLDelPayment, strSQLDelPaymentField, objRS, objConn	
			
		strSQLDelPaymentField = "DELETE FROM payment_field WHERE id_payment=?;"
		strSQLDelPayment = "DELETE FROM payment_type WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	

		Dim objCommand, objCommand2
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQLDelPaymentField
		objCommand2.CommandText = strSQLDelPayment
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)

		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand.Execute()
		end if
		objCommand2.Execute()
		
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = doubleValue
		
		'if (Application("dbType") = 0) then
			convertDoubleDelimiter = Replace(convertDoubleDelimiter, ".",",")
		'else		
			'convertDoubleDelimiter = Replace(convertDoubleDelimiter, ",",".")
		'end if			
	End Function

End Class
%>