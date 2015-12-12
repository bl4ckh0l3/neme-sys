<%

Class CardClass
	Private idCarrello	
	Private idUtente
	Private dtaCreazione
	
	
	Public Function getIDCarrello()
		getIDCarrello = idCarrello
	End Function
		
	Public Function getIDUtente()
		getIDUtente = idUtente
	End Function
	
	Public Function getDtaCreazione()
		getDtaCreazione = dtaCreazione
	End Function	
		
			
	Public Sub setIDCarrello(strIDCarrello)
		idCarrello = strIDCarrello
	End Sub		
			
	Public Sub setIDUtente(numIDUtente)
		idUtente = numIDUtente
	End Sub
	
	Public Sub setDtaCrezione(dtaCrezione)
		dtaCreazione = dtaCrezione
	End Sub

	
		
	Public Function getCarrelloByIDUser(id_user)
		on error resume next
		Dim objDB, strSQLInsert, strSQLSelect, objRS, objConn, objCarrello
		
		getCarrelloByIDUser = null

		strSQLSelect = "SELECT * FROM carrello WHERE id_utente=? ORDER BY dta_creazione DESC LIMIT 1;"

		Dim dta_ins, DD, MM, YY, HH, MIN, SS
		dta_ins = Now()	
		strSQLInsert = "INSERT INTO carrello(id_utente, dta_creazione) VALUES("
		strSQLInsert = strSQLInsert & "?, ?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user)
		Set objRS = objCommand.Execute()
		
		if objRS.EOF then
			Dim objCommand2
			Set objCommand2 = Server.CreateObject("ADODB.Command")
			objCommand2.ActiveConnection = objConn
			objCommand2.CommandType=1
			objCommand2.CommandText = strSQLInsert
			objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id_user)
			objCommand2.Parameters.Append objCommand2.CreateParameter(,135,1,,dta_ins)
			objCommand2.Execute()

			Set objRS = objCommand.Execute()
			if objRS.EOF then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=022")	
			else				
				Set objCarrello = New CardClass
				objCarrello.setIDCarrello(objRS("id_carrello"))
				objCarrello.setIDUtente(objRS("id_utente"))
				objCarrello.setDtaCrezione(objRS("dta_creazione"))													
				Set getCarrelloByIDUser = objCarrello
				Set objCarrello = nothing
			end if
		else
			Set objCarrello = New CardClass
			objCarrello.setIDCarrello(objRS("id_carrello"))
			objCarrello.setIDUtente(objRS("id_utente"))
			objCarrello.setDtaCrezione(objRS("dta_creazione"))						
			Set getCarrelloByIDUser = objCarrello
			Set objCarrello = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function
		
	Public Function findCarrelloByIDUser(id_user)
		on error resume next
		Dim objDB, strSQLInsert, strSQLSelect, objRS, objConn, objCarrello
		
		findCarrelloByIDUser = false
		
		strSQLSelect = "SELECT id_carrello FROM carrello WHERE id_utente=? ORDER BY dta_creazione DESC LIMIT 1;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user)
		Set objRS = objCommand.Execute()				
		
		if not (objRS.EOF) then
			findCarrelloByIDUser = true
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function
	
	Public Function getCarrelloByIDCarello(id_carrello)
		on error resume next
		Dim objDB, strSQLInsert, strSQLSelect, objRS, objConn, objCarrello
		
		getCarrelloByIDCarello = null
		
		strSQLSelect = "SELECT * FROM carrello WHERE id_carrello=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_carrello)
		Set objRS = objCommand.Execute()				
		
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=022")	
		else
			Set objCarrello = New CardClass
			objCarrello.setIDCarrello(objRS("id_carrello"))
			objCarrello.setIDUtente(objRS("id_utente"))
			objCarrello.setDtaCrezione(objRS("dta_creazione"))			
			Set getCarrelloByIDCarello = objCarrello
			Set objCarrello = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function
		
	Public Function getListaCarrelli()
		on error resume next
		Dim objDB, strSQLInsert, strSQLSelect, objRS, objConn, objCarrello
		
		getListaCarrelli = null
		
		strSQLSelect = "SELECT * FROM carrello ORDER BY dta_creazione DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
				
		Set objRS = objConn.Execute(strSQLSelect)
		
		if not objRS.EOF then
			Dim objDict, strID
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objCarrello = New CardClass
				strID = objRS("id_carrello")
				objCarrello.setIDCarrello(strID)
				objCarrello.setIDUtente(objRS("id_utente"))
				objCarrello.setDtaCrezione(objRS("dta_creazione"))
				objDict.add strID, objCarrello
				Set objCarrello = nothing
				objRS.moveNext()
			loop
			Set getListaCarrelli = objDict
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function
		
	public Sub updateIDUtenteCarrello(numIDCarrello, newUserID)
		on error resume next
		Dim objDB, objCarrello, objCurrCarrello, objRS, objConn	
			
		Set objCarrello = New CardClass
		Set objCurrCarrello = objCarrello.getCarrelloByIDCarello(numIDCarrello)
		
		if not(isNull(objCurrCarrello)) AND not(isEmpty(objCurrCarrello)) then	
			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()	
			
			strSQLUpdate = "UPDATE carrello SET id_utente=? WHERE id_carrello=?;"
			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQLUpdate
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,newUserID)
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDCarrello)
			objCommand.Execute()	
			Set objCommand = Nothing
			Set objDB = Nothing
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=019")
		end if	
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Sub

	Public Sub deleteCarrello(id_carrello)
		on error resume next
		Dim objDB, strSQLDelProdCard, strSQLDelProdFieldCard, strSQLDelCarr, objRS, objConn	
		strSQLDelProdCard = "DELETE FROM prodotti_x_carrello WHERE id_carrello=?;"
		strSQLDelProdFieldCard = "DELETE FROM product_fields_x_card WHERE id_card=?;"
		strSQLDelCarr = "DELETE FROM carrello WHERE id_carrello=?;"

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
		objCommand.CommandText = strSQLDelProdCard
		objCommand2.CommandText = strSQLDelProdFieldCard
		objCommand3.CommandText = strSQLDelCarr
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_carrello)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id_carrello)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id_carrello)		
		
		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand.Execute()
			objCommand2.Execute()
		end if
		objCommand3.Execute()
		
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
		
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	'public Sub toString()
		'response.write ()
	'end Sub
End Class
%>