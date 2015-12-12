<%

Class CommentsClass
	Private id_commento
	Private id_element
	Private element_type
	Private id_utente
	Private messaggio
	Private dta_ins
	Private vote_type
	Private active
	
	
	Public Function getIDCommento()
		getIDCommento = id_commento
	End Function
				
	Public Sub setIDCommento(numIDCommento)
		id_commento = numIDCommento
	End Sub
		
	Public Function getIDElement()
		getIDElement = id_element
	End Function
					
	Public Sub setIDElement(numIDElement)
		id_element = numIDElement
	End Sub
		
	Public Function getElementType()
		getElementType = element_type
	End Function
		
	Public Sub setElementType(strElementType)
		element_type = strElementType
	End Sub
		
	Public Function getIDUtente()
		getIDUtente = id_utente
	End Function
					
	Public Sub setIDUtente(numIDUtente)
		id_utente = numIDUtente
	End Sub
		
	Public Function getMessage()
		getMessage = messaggio
	End Function
		
	Public Sub setMessage(strMessage)
		messaggio = strMessage
	End Sub
		
	Public Function getDtaInserimento()
		getDtaInserimento = dta_ins
	End Function
		
	Public Sub setDtaInserimento(dtaInserimento)
		dta_ins = dtaInserimento
	End Sub
		
	Public Function getVoteType()
		getVoteType = vote_type
	End Function
		
	Public Sub setVoteType(strVoteType)
		vote_type = strVoteType
	End Sub
		
	Public Function getActive()
		getActive = active
	End Function
		
	Public Sub setActive(strActive)
		active = strActive
	End Sub
	


'*********************************** METODI PRODOTTO *********************** 				
	Public Function insertCommento(id_element, element_type, id_utente, strMessage, voteType, active, objConn)
		on error resume next
		Dim objDB, strSQL, strSQLSelect, objRS
		Dim dta_ins, DD, MM, YY, HH, MIN, SS
		dta_ins = Now()
		
		insertCommento = -1
				
		strSQL = "INSERT INTO commenti(id_element, element_type, id_utente, message, dta_inserimento, vote_type, active) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?);"
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,element_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMessage)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,voteType)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(commenti.id_commento) as id FROM commenti")
		if not (objRS.EOF) then
			insertCommento = objRS("id")	
		end if			
		Set objRS = Nothing		
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function	
	
	Public Function insertCommentoNoTransaction(id_element, element_type, id_utente, strMessage, voteType, active)
		on error resume next
		Dim objDB, strSQL, strSQLSelect, objRS, objConn
		Dim dta_ins, DD, MM, YY, HH, MIN, SS
		dta_ins = Now()
		
		insertCommentoNoTransaction = -1
				
		strSQL = "INSERT INTO commenti(id_element, element_type, id_utente, message, dta_inserimento, vote_type, active) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?);"
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,element_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMessage)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,voteType)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(commenti.id_commento) as id FROM commenti")
		if not (objRS.EOF) then
			insertCommentoNoTransaction = objRS("id")	
		end if			
		Set objRS = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyCommento(id_commento, id_element, element_type, id_utente, strMessage, voteType, active)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE commenti SET "	 
		strSQL = strSQL & "message=?,"
		strSQL = strSQL & "vote_type=?,"
		strSQL = strSQL & "active=?"
		strSQL = strSQL & " WHERE id_commento=? AND id_element=? AND id_utente=? AND element_type=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMessage)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,voteType)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_commento)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,element_type)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub updateStatus(id_commento, active)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE commenti SET "	 
		strSQL = strSQL & "active=?"
		strSQL = strSQL & " WHERE id_commento=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_commento)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteCommento(id)
		on error resume next
		Dim objDB, strSQLDelCommento, objRS, objConn
		strSQLDelCommento = "DELETE FROM commenti WHERE id_commento=?;" 		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDelCommento
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub				
		
	Public Sub deleteCommentiByIDElement(id, element_type)
		on error resume next
		Dim objDB, strSQLDelCommento, objRS, objConn
		strSQLDelCommento = "DELETE FROM commenti WHERE id_element=? AND element_type=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDelCommento
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,element_type)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
				
	Public Function findCommentiByType(element_type, active)
		on error resume next
		
		findCommentiByType = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM commenti WHERE element_type=?" 
		
		if not(isNull(active)) then strSQL = strSQL & " AND active=?"
		
		strSQL = strSQL & " ORDER BY id_element ASC, dta_inserimento DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,element_type)
		if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then						
			Dim objCommentoTmp, objDict, strIDTmp				
			Set objDict = Server.CreateObject("Scripting.Dictionary") 

					
			do while not objRS.EOF
				Set objCommentoTmp = New CommentsClass
				strIDTmp = objRS("id_commento")
				objCommentoTmp.setIDCommento(strIDTmp)
				objCommentoTmp.setIDElement(objRS("id_element")) 
				objCommentoTmp.setElementType(objRS("element_type")) 
				objCommentoTmp.setIDUtente(objRS("id_utente")) 
				objCommentoTmp.setMessage(objRS("message"))
				objCommentoTmp.setDtaInserimento(objRS("dta_inserimento"))	
				objCommentoTmp.setVoteType(objRS("vote_type"))
				objCommentoTmp.setActive(objRS("active"))
				objDict.add strIDTmp, objCommentoTmp
				Set objCommentoTmp = Nothing
				objRS.moveNext()
			loop
						
			Set findCommentiByType = objDict
			Set objDict = nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function		
				
	Public Function findCommentiByIDElement(id_element, element_type, active)
		on error resume next
		
		findCommentiByIDElement = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM commenti WHERE id_element=? AND element_type=?" 
		
		if not(isNull(active)) then strSQL = strSQL & " AND active=?"
		
		strSQL = strSQL & " ORDER BY dta_inserimento DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,element_type)
		if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then						
			Dim objCommentoTmp, objDict, strIDTmp				
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
					
			do while not objRS.EOF
				Set objCommentoTmp = New CommentsClass
				strIDTmp = objRS("id_commento")
				objCommentoTmp.setIDCommento(strIDTmp)
				objCommentoTmp.setIDElement(objRS("id_element")) 
				objCommentoTmp.setElementType(objRS("element_type")) 
				objCommentoTmp.setIDUtente(objRS("id_utente")) 
				objCommentoTmp.setMessage(objRS("message"))
				objCommentoTmp.setDtaInserimento(objRS("dta_inserimento"))	
				objCommentoTmp.setVoteType(objRS("vote_type"))
				objCommentoTmp.setActive(objRS("active"))
				objDict.add strIDTmp, objCommentoTmp
				Set objCommentoTmp = Nothing
				objRS.moveNext()
			loop
						
			Set findCommentiByIDElement = objDict
			Set objDict = nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function findCommentiByIDUtente(id_utente, id_element, element_type, active)
		on error resume next
		
		findCommentiByIDUtente = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM commenti WHERE id_utente=?"

		if (isNull(id_element) AND isNull(element_type) AND isNull(active)) then
			strSQL = "SELECT * FROM commenti WHERE id_utente=?"
		else
			if not(isNull(id_element)) then strSQL = strSQL & " AND id_element=?"
			if not(isNull(element_type)) then strSQL = strSQL & " AND element_type=?"
			if not(isNull(active)) then strSQL = strSQL & " AND active=?"
		end if
		strSQL = strSQL&" ORDER BY dta_inserimento DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_utente)
		if (isNull(id_element) AND isNull(element_type) AND isNull(active)) then
		else
			if not(isNull(id_element)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_element)
			if not(isNull(element_type)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,element_type)
			if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		end if		
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then				
			Dim objCommentoTmp, objDict, strIDTmp				
			Set objDict = Server.CreateObject("Scripting.Dictionary")
					
			do while not objRS.EOF
				Set objCommentoTmp = New CommentsClass
				strIDTmp = objRS("id_commento")
				objCommentoTmp.setIDCommento(strIDTmp)
				objCommentoTmp.setIDElement(objRS("id_element")) 
				objCommentoTmp.setElementType(objRS("element_type")) 
				objCommentoTmp.setIDUtente(objRS("id_utente")) 
				objCommentoTmp.setMessage(objRS("message"))
				objCommentoTmp.setDtaInserimento(objRS("dta_inserimento"))	
				objCommentoTmp.setVoteType(objRS("vote_type"))
				objCommentoTmp.setActive(objRS("active"))
				objDict.add strIDTmp, objCommentoTmp
				Set objCommentoTmp = Nothing
				objRS.moveNext()
			loop
						
			Set findCommentiByIDUtente = objDict
			Set objDict = nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function findCommentiByIDCommento(id_commento, element_type, active)
		on error resume next
		
		findCommentiByIDCommento = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM commenti WHERE id_commento=?" 
		
		if not(isNull(element_type)) then strSQL = strSQL & " AND element_type=?"
		if not(isNull(active)) then strSQL = strSQL & " AND active=?"
		
		strSQL = strSQL & ";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_commento)
		if not(isNull(element_type)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,element_type)
		if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then				

			Set objCommentoTmp = New CommentsClass
			strIDTmp = objRS("id_commento")
			objCommentoTmp.setIDCommento(strIDTmp)
			objCommentoTmp.setIDElement(objRS("id_element")) 
			objCommentoTmp.setElementType(objRS("element_type")) 
			objCommentoTmp.setIDUtente(objRS("id_utente")) 
			objCommentoTmp.setMessage(objRS("message"))
			objCommentoTmp.setDtaInserimento(objRS("dta_inserimento"))	
			objCommentoTmp.setVoteType(objRS("vote_type"))
			objCommentoTmp.setActive(objRS("active"))
						
			Set findCommentiByIDCommento = objCommentoTmp
			Set objDict = nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function countCommentiByIDUtente(id_utente, element_type, active)
		on error resume next
		
		countCommentiByIDUtente = 0
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT count(*) AS counter FROM commenti WHERE id_utente=?"

		if not(isNull(element_type)) then strSQL = strSQL & " AND element_type=?"
		if not(isNull(active)) then strSQL = strSQL & " AND active=?"
		strSQL = strSQL & ";"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_utente)
		if  not(isNull(element_type)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,element_type)
		if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then							
			countCommentiByIDUtente = (countCommentiByIDUtente + Cint(objRS("counter")))
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function countDistinctCommentiByIDUtente(id_utente, element_type, active)
		on error resume next
		
		countDistinctCommentiByIDUtente = 0
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT count(distinct(id_element)) AS counter FROM commenti WHERE id_utente=?"

		if not(isNull(element_type)) then strSQL = strSQL & " AND element_type=?"
		if not(isNull(active)) then strSQL = strSQL & " AND active=?"
		strSQL = strSQL & ";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_utente)
		if not(isNull(element_type)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,element_type)
		if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,active)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then							
			countDistinctCommentiByIDUtente = (objRS("counter"))
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	'public Sub toString()
		'response.write ()
	'end Sub
End Class
%>