<%
Class Bills4OrderClass

	private id_ordine
	private id_spesa
	private imponibile
	private tasse
	private totale
	private desc_spesa
	
	Public Function getIDOrdine()
		getIDOrdine = id_ordine
	End Function
				
	Public Sub setIDOrdine(numIDOrdine)
		id_ordine = numIDOrdine
	End Sub
	
	Public Function getIDSpesa()
		getIDSpesa = id_spesa
	End Function
				
	Public Sub setIDSpesa(numIDSpesa)
		id_spesa = numIDSpesa
	End Sub
	
	Public Function getImponibile()
		getImponibile = Cdbl(imponibile)
	End Function
				
	Public Sub setImponibile(numImponibile)
		imponibile = numImponibile
	End Sub
	
	Public Function getTasse()
		getTasse = Cdbl(tasse)
	End Function
				
	Public Sub setTasse(numTasse)
		tasse = numTasse
	End Sub
	
	Public Function getTotale()
		getTotale = Cdbl(totale)
	End Function
				
	Public Sub setTotale(numTotale)
		totale = numTotale
	End Sub
	
	Public Function getDescSpesa()
		getDescSpesa = desc_spesa
	End Function
				
	Public Sub setDescSpesa(strDescSpesa)
		desc_spesa = strDescSpesa
	End Sub
	

	'*************************** METODI SPESE PER ORDINE ***********************
	
	Public Function getSpeseXOrdine(id_ordine)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objOrdine, objSpeseXOrdine
		getSpeseXOrdine = null  
		strSQL = "SELECT * FROM spese_x_ordine WHERE id_ordine=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ordine)
		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then  
			Set objDict = Server.CreateObject("Scripting.Dictionary")   
			do while not objRS.EOF
				Set objSpeseXOrdine = new Bills4OrderClass
				strID = objRS("id_spesa")
				objSpeseXOrdine.setIDordine(objRS("id_ordine"))
				objSpeseXOrdine.setIDSpesa(strID)
				objSpeseXOrdine.setImponibile(objRS("imponibile"))
				objSpeseXOrdine.setTasse(objRS("tasse"))
				objSpeseXOrdine.setTotale(objRS("totale"))
				objSpeseXOrdine.setDescSpesa(objRS("desc_spesa"))
				
				objDict.add strID, objSpeseXOrdine
				Set objSpeseXOrderClass = nothing
				objRS.moveNext()
			loop
			
			Set getSpeseXOrdine = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
	
	Public Sub insertSpeseXOrdineNoTransaction(id_ordine, id_spesa, imponibile, tasse, totale, desc_spesa)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		
		strSQL = "INSERT INTO spese_x_ordine(id_ordine, id_spesa, imponibile, tasse, totale, desc_spesa) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"
						
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(imponibile))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(tasse))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(totale))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,desc_spesa)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifySpeseXOrdineNoTransaction(id_ordine, id_spesa, imponibile, tasse, totale, desc_spesa)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		strSQL = "UPDATE spese_x_ordine SET "
		strSQL = strSQL & "id_ordine=?,"
		strSQL = strSQL & "id_spesa=?,"
		strSQL = strSQL & "imponibile=?,"
		strSQL = strSQL & "tasse=?,"
		strSQL = strSQL & "totale=?,"
		strSQL = strSQL & "desc_spesa=?"
		strSQL = strSQL & " WHERE id_ordine=? AND id_spesa=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(imponibile))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(tasse))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(totale))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,desc_spesa)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
	
	Public Sub insertSpeseXOrdine(id_ordine, id_spesa, imponibile, tasse, totale, desc_spesa, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO spese_x_ordine(id_ordine, id_spesa, imponibile, tasse, totale, desc_spesa) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(imponibile))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(tasse))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(totale))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,desc_spesa)
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
		
	Public Sub modifySpeseXOrdine(id_ordine, id_spesa, imponibile, tasse, totale, desc_spesa, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE spese_x_ordine SET "
		strSQL = strSQL & "id_ordine=?,"
		strSQL = strSQL & "id_spesa=?,"
		strSQL = strSQL & "imponibile=?,"
		strSQL = strSQL & "tasse=?,"
		strSQL = strSQL & "totale=?,"
		strSQL = strSQL & "desc_spesa=?"
		strSQL = strSQL & " WHERE id_ordine=? AND id_spesa=?;" 

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(imponibile))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(tasse))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(totale))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,desc_spesa)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
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
		
	Public Sub deleteSpeseXOrdineNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn 
		strSQL = "DELETE FROM spese_x_ordine WHERE id_ordine=?;"

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
		
	Public Sub deleteSpeseXOrdine(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM spese_x_ordine WHERE id_ordine=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
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
	
	Public Sub deleteSpesaXOrdineNoTransaction(id_ordine, id_spesa)
		on error resume next
		Dim objDB, strSQL, objRS, objConn 
		strSQL = "DELETE FROM spese_x_ordine WHERE id_ordine=? AND id_spesa=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub deleteSpesaXOrdine(id_ordine, id_spesa, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM spese_x_ordine WHERE id_ordine=? AND id_spesa=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
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