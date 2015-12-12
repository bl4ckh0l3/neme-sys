<%

Class TaxsClass
	
	Private id
	Private descrizione_tassa
	Private valore
	Private tipologia_valore
	
	
	Public Function getTasseID()
		getTasseID = id
	End Function
	
	Public Sub setTasseID(strID)
		id = strID
	End Sub	
	
	Public Function getDescrizioneTassa()
		getDescrizioneTassa = descrizione_tassa
	End Function
	
	Public Sub setDescrizioneTassa(strDesc)
		descrizione_tassa = strDesc
	End Sub		
	
	Public Function getValore()
		getValore = Cdbl(valore)
	End Function
	
	Public Sub setValore(strValore)
		valore = strValore
	End Sub
	
	
	Public Function getTipoValore()
		getTipoValore = tipologia_valore
	End Function
	
	Public Sub setTipoValore(strTipoValore)
		tipologia_valore = strTipoValore
	End Sub
	
	
	Public Function getListaTasse(descrizione_tassa, tipologia_valore)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getListaTasse = null		
		strSQL = "SELECT * FROM tasse"
		
		if (isNull(descrizione_tassa) AND isNull(tipologia_valore)) then
			strSQL = "SELECT * FROM tasse"
		else
			strSQL = strSQL & " WHERE"
			if not(isNull(descrizione_tassa)) then strSQL = strSQL & " AND descrizione_tassa=?"
			if not(isNull(tipologia_valore)) then strSQL = strSQL & " AND tipologia_valore=?"
		end if
		
		strSQL = strSQL & " ORDER BY descrizione_tassa DESC"
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = strSQL & ";"
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		if (isNull(descrizione_tassa) AND isNull(tipologia_valore)) then
		else
			if not(isNull(descrizione_tassa)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,descrizione_tassa)
			if not(isNull(tipologia_valore)) then objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,tipologia_valore)
		end if		

		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objTasse
			do while not objRS.EOF				
				Set objTasse = new TaxsClass
				strID = objRS("id")
				objTasse.setTasseID(strID)
				objTasse.setDescrizioneTassa(objRS("descrizione_tassa"))
				objTasse.setValore(objRS("valore"))	
				objTasse.setTipoValore(objRS("tipologia_valore"))	
				objDict.add strID, objTasse
				objRS.moveNext()
			loop
			Set objTasse = nothing							
			Set getListaTasse = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	
	Public Function findTassaByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		findTassaByID = null		
		strSQL = "SELECT * FROM tasse WHERE id =?;"
		
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
			Dim objTasse		
			Set objTasse = new TaxsClass
			strID = objRS("id")
			objTasse.setTasseID(strID)
			objTasse.setDescrizioneTassa(objRS("descrizione_tassa"))
			objTasse.setValore(objRS("valore"))	
			objTasse.setTipoValore(objRS("tipologia_valore"))					
			Set findTassaByID = objTasse			
			Set objTasse = nothing			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Sub insertTassa(descrizione_tassa, valore, tipologia_valore)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		
		strSQL = "INSERT INTO tasse(descrizione_tassa, valore, tipologia_valore) VALUES("
		strSQL = strSQL & "?,?,?);"
						
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,descrizione_tassa)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,tipologia_valore)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyTassa(id, descrizione_tassa, valore, tipologia_valore)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		strSQL = "UPDATE tasse SET "
		strSQL = strSQL & "id=?,"
		strSQL = strSQL & "descrizione_tassa=?,"
		strSQL = strSQL & "valore=?,"
		strSQL = strSQL & "tipologia_valore=?"
		strSQL = strSQL & " WHERE id=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,descrizione_tassa)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,tipologia_valore)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteTassa(id)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM tasse WHERE id=?;"
		
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

	Public Function findTasseAssociations(id_tassa)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		findTasseAssociations = false	
		strSQL = "SELECT prodotti.id_tassa_applicata FROM prodotti WHERE prodotti.id_tassa_applicata=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_tassa)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then							
			findTasseAssociations = true				
		end if		
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

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