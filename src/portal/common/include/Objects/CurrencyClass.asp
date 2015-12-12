<%
Class CurrencyClass
	
	Private id
	Private currencyKey
	Private rate
	Private dta_riferimento
	Private dta_inserimento
	Private active
	Private isDefault
	
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub	
	
	Public Function getCurrency()
		getCurrency = currencyKey
	End Function
	
	Public Sub setCurrency(strCurr)
		currencyKey = strCurr
	End Sub		
	
	Public Function getRate()
		getRate = rate
	End Function
	
	Public Sub setRate(strRate)
		rate = strRate
	End Sub	
	
	Public Function getDtaRefer()
		getDtaRefer = dta_riferimento
	End Function
	
	Public Sub setDtaRefer(strTimeRef)
		dta_riferimento = strTimeRef
	End Sub	
	
	Public Function getDtaInsert()
		getDtaInsert = dta_inserimento
	End Function
	
	Public Sub setDtaInsert(strTime)
		dta_inserimento = strTime
	End Sub	
	
	Public Function getActive()
		getActive = active
	End Function
	
	Public Sub setActive(strActive)
		active = strActive
	End Sub
	
	Public Function getDefault()
		getDefault = isDefault
	End Function
	
	Public Sub setDefault(strDefault)
		isDefault = strDefault
	End Sub
	
	
	Public Function getListaCurrency(currencyKey, active, isDefault)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getListaCurrency = null		
		strSQL = "SELECT * FROM currency"
		
		if (isNull(currencyKey) AND isNull(active) AND isNull(isDefault)) then
			strSQL = "SELECT * FROM currency"
		else
			strSQL = strSQL & " WHERE"
			if not(isNull(currencyKey)) then strSQL = strSQL & " AND currency=?"
			if not(isNull(active)) then strSQL = strSQL & " AND active=?"
			if not(isNull(isDefault)) then strSQL = strSQL & " AND is_default=?"
		end if
		
		strSQL = strSQL & " ORDER BY currency ASC;"
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if (isNull(currencyKey) AND isNull(active) AND isNull(isDefault)) then
		else
			if not(isNull(currencyKey)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,5,currencyKey)
			if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active)
			if not(isNull(isDefault)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,isDefault)			
		end if
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objCurrency
			do while not objRS.EOF				
				Set objCurrency = new CurrencyClass
				strID = objRS("id")
				objCurrency.setID(strID)
				objCurrency.setCurrency(objRS("currency"))
				objCurrency.setRate(objRS("rate"))	
				objCurrency.setDtaInsert(objRS("dta_inserimento"))
				objCurrency.setDtaRefer(objRS("dta_riferimento"))	
				objCurrency.setActive(objRS("active"))
				objCurrency.setDefault(objRS("is_default"))
				objDict.add strID, objCurrency
				objRS.moveNext()
			loop
			Set objCurrency = nothing							
			Set getListaCurrency = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findCurrencyByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		findCurrencyByID = null		
		strSQL = "SELECT * FROM currency WHERE id =?;"
		strSQL = Trim(strSQL)
		
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
			Dim objCurrency		
			Set objCurrency = new CurrencyClass
			strID = objRS("id")
			objCurrency.setID(strID)
			objCurrency.setCurrency(objRS("currency"))
			objCurrency.setRate(objRS("rate"))	
			objCurrency.setDtaInsert(objRS("dta_inserimento"))	
			objCurrency.setDtaRefer(objRS("dta_riferimento"))
			objCurrency.setActive(objRS("active"))	
			objCurrency.setDefault(objRS("is_default"))		
			Set findCurrencyByID = objCurrency			
			Set objCurrency = nothing			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	
	Public Function findCurrencyByCurrency(currencyKey)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		findCurrencyByCurrency = null		
		strSQL = "SELECT * FROM currency WHERE currency =?;"
		strSQL = Trim(strSQL)
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,5,currencyKey)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Dim objCurrency		
			Set objCurrency = new CurrencyClass
			strID = objRS("id")
			objCurrency.setID(strID)
			objCurrency.setCurrency(objRS("currency"))
			objCurrency.setRate(objRS("rate"))	
			objCurrency.setDtaInsert(objRS("dta_inserimento"))	
			objCurrency.setDtaRefer(objRS("dta_riferimento"))
			objCurrency.setActive(objRS("active"))	
			objCurrency.setDefault(objRS("is_default"))			
			Set findCurrencyByCurrency = objCurrency			
			Set objCurrency = nothing			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	
	Public Function getDefaultCurrency()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getDefaultCurrency = null		
		strSQL = "SELECT * FROM currency WHERE is_default =1;"
		strSQL = Trim(strSQL)
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Dim objCurrency		
			Set objCurrency = new CurrencyClass
			strID = objRS("id")
			objCurrency.setID(strID)
			objCurrency.setCurrency(objRS("currency"))
			objCurrency.setRate(objRS("rate"))	
			objCurrency.setDtaInsert(objRS("dta_inserimento"))	
			objCurrency.setDtaRefer(objRS("dta_riferimento"))
			objCurrency.setActive(objRS("active"))	
			objCurrency.setDefault(objRS("is_default"))			
			Set getDefaultCurrency = objCurrency			
			Set objCurrency = nothing			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function convertCurrency(amount, currFrom, currTo)
		on error resume next
		Dim objDB, strSQL1,strSQL2, objRS1,objRS2, objConn
		Dim rate1, rate2

		convertCurrency = null		

		strSQL = "SELECT c1.rate AS rate1, c2.rate AS rate2 FROM currency AS c1 INNER JOIN currency AS c2 WHERE c1.currency = ? AND c2.currency = ?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,5,currFrom)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,5,currTo)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then	
			rate1 = objRS("rate1")	
			rate2 = objRS("rate2")
			convertCurrency = Cdbl(amount)*(CDbl(rate2)/CDbl(rate1))		
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	
	Public Sub insertCurrency(currencyKey, rate, dta_riferimento, dta_inserimento, active, isDefault)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		
		strSQL = "INSERT INTO currency(currency, rate, dta_riferimento, dta_inserimento, active, is_default) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,5,currencyKey)		
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate))
		objCommand.Parameters.Append objCommand.CreateParameter(,133,1,,dta_riferimento)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_inserimento)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,isDefault)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyCurrency(id, currencyKey, rate, dta_riferimento, dta_inserimento, active, isDefault)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		strSQL = "UPDATE currency SET "
		strSQL = strSQL & "id=?,"
		strSQL = strSQL & "currency=?,"
		strSQL = strSQL & "rate=?,"
		strSQL = strSQL & "dta_riferimento=?,"
		strSQL = strSQL & "dta_inserimento=?,"
		strSQL = strSQL & "active=?,"
		strSQL = strSQL & "is_default=?"
		strSQL = strSQL & " WHERE id=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,5,currencyKey)		
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate))		
		objCommand.Parameters.Append objCommand.CreateParameter(,133,1,,dta_riferimento)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_inserimento)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,isDefault)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub resetDefaultCurrency()
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		strSQL = "UPDATE currency SET "
		strSQL = strSQL & "is_default=0;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub deleteCurrency(id)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM currency WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		Set objRS = objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteAllCurrency()
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM currency;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Set objRS = objConn.Execute(strSQL)
		Set objRS = Nothing
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