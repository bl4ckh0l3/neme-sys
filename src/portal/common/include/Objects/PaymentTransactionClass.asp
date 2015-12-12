<%
Class PaymentTransactionClass
	Private id
	Private idOrdine
	Private idModulo
	Private idTransaction
	Private paymentStatus
	Private notified
	private insertDate
	
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub
	
	Public Function getIdOrdine()
		getIdOrdine = idOrdine
	End Function
	
	Public Sub setIdOrdine(strIdOrdine)
		idOrdine = strIdOrdine
	End Sub
	
	Public Function getIdModulo()
		getIdModulo = idModulo
	End Function
	
	Public Sub setIdModulo(strIdModulo)
		idModulo = strIdModulo
	End Sub
	
	Public Function getIdTransaction()
		getIdTransaction = idTransaction
	End Function
	
	Public Sub setIdTransaction(strIdTransaction)
		idTransaction = strIdTransaction
	End Sub
	
	Public Function getPaymentStatus()
		getPaymentStatus = paymentStatus
	End Function
	
	Public Sub setPaymentStatus(strPaymentStatus)
		paymentStatus = strPaymentStatus
	End Sub
	
	Public Function isNotified()
		isNotified = notified
	End Function
	
	Public Sub setNotified(bolNotified)
		notified = bolNotified
	End Sub
	
	Public Function getInsertDate()
		getInsertDate = insertDate
	End Function
	
	Public Sub setInsertDate(datInsertDate)
		insertDate = datInsertDate
	End Sub
	
		
	Public Function getListaOrderPaymentTransaction(idOrdine)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaOrderPaymentTransaction = null		
		strSQL = "SELECT * FROM payment_transactions WHERE id_ordine=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		Set objRS = objCommand.Execute()			
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objPaymentTRansaction
			do while not objRS.EOF				
				Set objPaymentTRansaction = new PaymentTransactionClass
				strID = objRS("id")
				objPaymentTRansaction.setID(strID)
				objPaymentTRansaction.setIdOrdine(objRS("id_ordine"))	
				objPaymentTRansaction.setIdModulo(objRS("id_modulo"))	
				objPaymentTRansaction.setIdTransaction(objRS("id_transaction"))	
				objPaymentTRansaction.setPaymentStatus(objRS("status"))	
				objPaymentTRansaction.setNotified(objRS("notified"))	
				objPaymentTRansaction.setInsertDate(objRS("insert_date"))
				objDict.add strID, objPaymentTRansaction
				objRS.moveNext()
			loop
			Set objPaymentTRansaction = nothing							
			Set getListaOrderPaymentTransaction = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findPaymentTransactionByIDTransaction(idOrdine, idTransaction)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentTransactionByIDTransaction = null		
		strSQL = "SELECT * FROM payment_transactions WHERE id_ordine=? AND  id_transaction=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idTransaction)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objPaymentTRansaction
			Set objPaymentTRansaction = new PaymentTransactionClass
			objPaymentTRansaction.setID(objRS("id"))
			objPaymentTRansaction.setIdOrdine(objRS("id_ordine"))	
			objPaymentTRansaction.setIdModulo(objRS("id_modulo"))	
			objPaymentTRansaction.setIdTransaction(objRS("id_transaction"))	
			objPaymentTRansaction.setPaymentStatus(objRS("status"))	
			objPaymentTRansaction.setNotified(objRS("notified"))
			objPaymentTRansaction.setInsertDate(objRS("insert_date"))
			Set findPaymentTransactionByIDTransaction = objPaymentTRansaction
			Set objPaymentTRansaction = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findPaymentTransactionByIDTransactionToNotify(idOrdine)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentTransactionByIDTransactionToNotify = null
		Set objUtil = new UtilClass		
		strSQL = "SELECT * FROM payment_transactions WHERE id_ordine=? AND status=? AND notified=0 ORDER BY id DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,objUtil.getUniqueKeySuccessPaymentTransaction())
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objPaymentTRansaction
			Set objPaymentTRansaction = new PaymentTransactionClass
			objPaymentTRansaction.setID(objRS("id"))
			objPaymentTRansaction.setIdOrdine(objRS("id_ordine"))
			objPaymentTRansaction.setIdModulo(objRS("id_modulo"))		
			objPaymentTRansaction.setIdTransaction(objRS("id_transaction"))	
			objPaymentTRansaction.setPaymentStatus(objRS("status"))	
			objPaymentTRansaction.setNotified(objRS("notified"))
			objPaymentTRansaction.setInsertDate(objRS("insert_date"))
			Set findPaymentTransactionByIDTransactionToNotify = objPaymentTRansaction			
			Set objPaymentTRansaction = nothing				
		end if
		
		Set objUtil = nothing
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findPaymentTransactionByIDTransactionNotified(idOrdine, idTransaction)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findPaymentTransactionByIDTransactionNotified = null
		Set objUtil = new UtilClass		
		strSQL = "SELECT * FROM payment_transactions WHERE id_ordine=? AND  id_transaction=? AND status=? AND notified=1;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idTransaction)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,objUtil.getUniqueKeySuccessPaymentTransaction())
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objPaymentTRansaction
			Set objPaymentTRansaction = new PaymentTransactionClass
			objPaymentTRansaction.setID(objRS("id"))
			objPaymentTRansaction.setIdOrdine(objRS("id_ordine"))
			objPaymentTRansaction.setIdModulo(objRS("id_modulo"))		
			objPaymentTRansaction.setIdTransaction(objRS("id_transaction"))	
			objPaymentTRansaction.setPaymentStatus(objRS("status"))	
			objPaymentTRansaction.setNotified(objRS("notified"))
			objPaymentTRansaction.setInsertDate(objRS("insert_date"))
			Set findPaymentTransactionByIDTransactionNotified = objPaymentTRansaction
			Set objPaymentTRansaction = nothing				
		end if
		
		Set objUtil = nothing
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function isPaymentTransactionNotified(idOrdine, idTransaction)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		isPaymentTransactionNotified = false		
		strSQL = "SELECT * FROM payment_transactions WHERE id_ordine=? AND  id_transaction=? AND notified=1;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idTransaction)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			isPaymentTransactionNotified = true				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function hasPaymentTransactionNotified(idOrdine)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		hasPaymentTransactionNotified = false		
		strSQL = "SELECT * FROM payment_transactions WHERE id_ordine=? AND notified=1;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		Set objRS = objCommand.Execute()				
		
		if not(objRS.EOF) then
			hasPaymentTransactionNotified = true				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertPaymentTransaction(idOrdine, idModulo, idTransaction, status, notified, insertDate, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		if (Application("dbType") = 1) then
			insertDate = convertDate(insertDate)
		end if	
		
		strSQL = "INSERT INTO payment_transactions(id_ordine, id_modulo, id_transaction, status, notified, insert_date) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idModulo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idTransaction)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,status)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,notified)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,insertDate)
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
			
	Public Sub insertPaymentTransactionNoTrans(idOrdine, idModulo, idTransaction, status, notified, insertDate)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		if (Application("dbType") = 1) then
			insertDate = convertDate(insertDate)
		end if	
		
		strSQL = "INSERT INTO payment_transactions(id_ordine, id_modulo, id_transaction, status, notified, insert_date) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idModulo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idTransaction)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,status)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,notified)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,insertDate)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyPaymentTransaction(id, idOrdine, idModulo, idTransaction, status, notified, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE payment_transactions SET "
		strSQL = strSQL & "id_ordine=?,"
		strSQL = strSQL & "id_modulo=?,"
		strSQL = strSQL & "id_transaction=?,"
		strSQL = strSQL & "status=?,"		
		strSQL = strSQL & "notified=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idModulo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idTransaction)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,status)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,notified)
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
		
	Public Sub modifyPaymentTransactionNoTrans(id, idOrdine, idModulo, idTransaction, status, notified)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE payment_transactions SET "
		strSQL = strSQL & "id_ordine=?,"
		strSQL = strSQL & "id_modulo=?,"
		strSQL = strSQL & "id_transaction=?,"
		strSQL = strSQL & "status=?,"		
		strSQL = strSQL & "notified=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idModulo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idTransaction)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,status)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,notified)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deletePaymentTransaction(idOrdine, idTransaction)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM payment_transactions WHERE id_ordine=? AND  id_transaction=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrdine)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,idTransaction)
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