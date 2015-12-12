<%

Class TargetClass
	Private id
	Private descrizione
	Private tipo
	Private locked
	Private automatic
	
	
	Public Function getTargetID()
		getTargetID = id
	End Function
	
	Public Sub setTargetID(strID)
		id = strID
	End Sub
	
	Public Function getTargetDescrizione()
		getTargetDescrizione = descrizione
	End Function
	
	Public Sub setTargetDescrizione(strDesc)
		descrizione = strDesc
	End Sub
	
	Public Function getTargetType()
		getTargetType = tipo
	End Function
	
	Public Sub setTargetType(strType)
		tipo = strType
	End Sub
	
	Public Function isLocked()
		isLocked = locked
	End Function
	
	Public Sub setLocked(intLocked)
		locked = intLocked
	End Sub
	
	Public Function isAutomatic()
		isAutomatic = automatic
	End Function
	
	Public Sub setAutomatic(intAutomatic)
		automatic = intAutomatic
	End Sub
	
		
	Public Function getListaTarget()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaTarget = null		
		strSQL = "SELECT * FROM target ORDER BY type, descrizione;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objTarget
			do while not objRS.EOF				
				Set objTarget = new TargetClass
				strID = objRS("id")
				objTarget.setTargetID(strID)
				objTarget.setTargetDescrizione(objRS("descrizione"))	
				objTarget.setTargetType(objRS("type"))	
				objTarget.setLocked(objRS("locked"))	
				objTarget.setAutomatic(objRS("automatic"))
				objDict.add strID, objTarget
				objRS.moveNext()
			loop
			Set objTarget = nothing							
			Set getListaTarget = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function getListLockedTarget()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListLockedTarget = null		
		strSQL = "SELECT * FROM target WHERE locked=1 ORDER BY type, descrizione;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		Set objRS = objCommand.Execute()
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objTarget
			do while not objRS.EOF				
				Set objTarget = new TargetClass
				strID = objRS("id")
				objTarget.setTargetID(strID)
				objTarget.setTargetDescrizione(objRS("descrizione"))	
				objTarget.setTargetType(objRS("type"))	
				objTarget.setLocked(objRS("locked"))	
				objTarget.setAutomatic(objRS("automatic"))
				objDict.add strID, objTarget
				objRS.moveNext()
			loop
			Set objTarget = nothing							
			Set getListLockedTarget = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function getListAutomaticTarget()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListAutomaticTarget = null		
		strSQL = "SELECT * FROM target WHERE automatic=1 ORDER BY type, descrizione;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		Set objRS = objCommand.Execute()
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objTarget
			do while not objRS.EOF				
				Set objTarget = new TargetClass
				strID = objRS("id")
				objTarget.setTargetID(strID)
				objTarget.setTargetDescrizione(objRS("descrizione"))	
				objTarget.setTargetType(objRS("type"))	
				objTarget.setLocked(objRS("locked"))	
				objTarget.setAutomatic(objRS("automatic"))
				objDict.add strID, objTarget
				objRS.moveNext()
			loop
			Set objTarget = nothing							
			Set getListAutomaticTarget = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findTarget(id_target)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findTarget = null		
		strSQL = "SELECT * FROM target WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_target)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objTarget
			Set objTarget = new TargetClass
			objTarget.setTargetID(objRS("id"))
			objTarget.setTargetDescrizione(objRS("descrizione"))	
			objTarget.setTargetType(objRS("type"))		
			objTarget.setLocked(objRS("locked"))
			objTarget.setAutomatic(objRS("automatic"))						
			Set findTarget = objTarget
			Set objTarget = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findTargetByID(id_target)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findTargetByID = null		
		strSQL = "SELECT * FROM target WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_target)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objTarget
			Set objTarget = new TargetClass
			objTarget.setTargetID(objRS("id"))
			objTarget.setTargetDescrizione(objRS("descrizione"))		
			objTarget.setTargetType(objRS("type"))			
			objTarget.setLocked(objRS("locked"))
			objTarget.setAutomatic(objRS("automatic"))								
			Set findTargetByID = objTarget
			Set objTarget = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function findTargetByDesc(desc_target, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, objDict
		findTargetByDesc = null		
		strSQL = "SELECT * FROM target WHERE descrizione LIKE ?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,"%"&desc_target&"%")
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objTarget
			Set objTarget = new TargetClass
			objTarget.setTargetID(objRS("id"))
			objTarget.setTargetDescrizione(objRS("descrizione"))	
			objTarget.setTargetType(objRS("type"))							
			objTarget.setLocked(objRS("locked"))	
			objTarget.setAutomatic(objRS("automatic"))					
			Set findTargetByDesc = objTarget
			Set objTarget = nothing				
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
		
	Public Function findTargetByDescNoTransaction(desc_target)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findTargetByDescNoTransaction = null		
		strSQL = "SELECT * FROM target WHERE descrizione LIKE ?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,"%"&desc_target&"%")
		Set objRS = objCommand.Execute()				
		
		if not(objRS.EOF) then
			Dim objTarget
			Set objTarget = new TargetClass
			objTarget.setTargetID(objRS("id"))
			objTarget.setTargetDescrizione(objRS("descrizione"))	
			objTarget.setTargetType(objRS("type"))							
			objTarget.setLocked(objRS("locked"))
			objTarget.setAutomatic(objRS("automatic"))						
			Set findTargetByDescNoTransaction = objTarget
			Set objTarget = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function findTargetByDescEq(desc_target, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, objDict
		findTargetByDescEq = null		
		strSQL = "SELECT * FROM target WHERE descrizione=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,desc_target)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objTarget
			Set objTarget = new TargetClass
			objTarget.setTargetID(objRS("id"))
			objTarget.setTargetDescrizione(objRS("descrizione"))	
			objTarget.setTargetType(objRS("type"))							
			objTarget.setLocked(objRS("locked"))
			objTarget.setAutomatic(objRS("automatic"))						
			Set findTargetByDescEq = objTarget
			Set objTarget = nothing				
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
		
	Public Function findTargetByDescEqNoTransaction(desc_target)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findTargetByDescEqNoTransaction = null		
		strSQL = "SELECT * FROM target WHERE descrizione=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,desc_target)
		Set objRS = objCommand.Execute()				
		
		if not(objRS.EOF) then
			Dim objTarget
			Set objTarget = new TargetClass
			objTarget.setTargetID(objRS("id"))
			objTarget.setTargetDescrizione(objRS("descrizione"))	
			objTarget.setTargetType(objRS("type"))							
			objTarget.setLocked(objRS("locked"))
			objTarget.setAutomatic(objRS("automatic"))							
			Set findTargetByDescEqNoTransaction = objTarget
			Set objTarget = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function		
		
	Public Function findTargetByDescAndType(desc_target, ttype, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, objDict
		findTargetByDescAndType = null		
		strSQL = "SELECT * FROM target WHERE descrizione LIKE ? AND type=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,"%"&desc_target&"%")
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,ttype)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objTarget
			Set objTarget = new TargetClass
			objTarget.setTargetID(objRS("id"))
			objTarget.setTargetDescrizione(objRS("descrizione"))	
			objTarget.setTargetType(objRS("type"))							
			objTarget.setLocked(objRS("locked"))
			objTarget.setAutomatic(objRS("automatic"))							
			Set findTargetByDescAndType = objTarget
			Set objTarget = nothing				
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
		
	Public Function findTargetByDescEqAndType(desc_target, ttype, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, objDict
		findTargetByDescEqAndType = null		
		strSQL = "SELECT * FROM target WHERE descrizione=? AND type=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,desc_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,ttype)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objTarget
			Set objTarget = new TargetClass
			objTarget.setTargetID(objRS("id"))
			objTarget.setTargetDescrizione(objRS("descrizione"))	
			objTarget.setTargetType(objRS("type"))							
			objTarget.setLocked(objRS("locked"))
			objTarget.setAutomatic(objRS("automatic"))							
			Set findTargetByDescEqAndType = objTarget
			Set objTarget = nothing				
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
		
	Public Function findTargetsByCategoria(id_cat)		
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findTargetsByCategoria = null		
		strSQL = "SELECT target_x_categoria.id_target, target.descrizione FROM target_x_categoria, target WHERE target_x_categoria.id_categoria=? AND target_x_categoria.id_target = target.id"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_cat)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objTarget
			do while not objRS.EOF
				strID = objRS("id_target")
				strDesc = objRS("descrizione")	
				
				Set objTarget = new TargetClass
				objTarget.setTargetID(strID)
				objTarget.setTargetDescrizione(strDesc)		
				objTarget.setTargetType(objRS("type"))								
				objTarget.setLocked(objRS("locked"))
				objTarget.setAutomatic(objRS("automatic"))												
				objDict.add strID, objTarget
				objRS.moveNext()
			loop			
			Set objTarget = nothing	
							
			Set findTargetsByCategoria = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function		
		
	Public Function findTargetsByType(ttype)		
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findTargetsByType = null		
		strSQL = "SELECT * FROM target WHERE type=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,ttype)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objTarget
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("descrizione")	
				
				Set objTarget = new TargetClass
				objTarget.setTargetID(strID)
				objTarget.setTargetDescrizione(strDesc)		
				objTarget.setTargetType(objRS("type"))								
				objTarget.setLocked(objRS("locked"))
				objTarget.setAutomatic(objRS("automatic"))											
				objDict.add strID, objTarget
				objRS.moveNext()
			loop			
			Set objTarget = nothing	
							
			Set findTargetsByType = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertTarget(strDescrizione, tipo, locked, automatic, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO target(descrizione, type, locked, automatic) VALUES("
		strSQL = strSQL & "?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,locked)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,automatic)
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
		
	Public Sub modifyTarget(id, strDescrizione, tipo, locked, automatic, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE target SET "
		strSQL = strSQL & "descrizione=?,"
		strSQL = strSQL & "type=?,"		
		strSQL = strSQL & "locked=?,"	
		strSQL = strSQL & "automatic=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,locked)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,automatic)
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
			
	Public Sub insertTargetNoTransaction(strDescrizione, tipo, locked, automatic)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO target(descrizione, type, locked, automatic) VALUES("
		strSQL = strSQL & "?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,locked)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,automatic)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyTargetNoTransaction(id, strDescrizione, tipo, locked, automatic)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE target SET "
		strSQL = strSQL & "descrizione=?,"
		strSQL = strSQL & "type=?,"		
		strSQL = strSQL & "locked=?,"	
		strSQL = strSQL & "automatic=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,locked)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,automatic)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteTarget(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM target WHERE id=?;"

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

	Public Function findTargetAssociations(id_target)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		Dim strSQL2, strSQL3, strSQL4
		findTargetAssociations = false	
		strSQL = "SELECT target_x_utente.id_target FROM target_x_utente WHERE target_x_utente.id_target=?;"
		strSQL2 = "SELECT target_x_news.id_target FROM target_x_news WHERE target_x_news.id_target=?;"
		strSQL3 = "SELECT target_x_categoria.id_target FROM target_x_categoria WHERE target_x_categoria.id_target=?;"
		strSQL4 = "SELECT target_x_prodotto.id_target FROM target_x_prodotto WHERE target_x_prodotto.id_target=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand, objCommand2, objCommand3, objCommand4
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		Set objCommand3 = Server.CreateObject("ADODB.Command")
		Set objCommand4 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand3.ActiveConnection = objConn
		objCommand4.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand3.CommandType=1
		objCommand4.CommandType=1
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQL2
		objCommand3.CommandText = strSQL3
		objCommand4.CommandText = strSQL4
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_target)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,3,1,,id_target)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,3,1,,id_target)
		objCommand4.Parameters.Append objCommand4.CreateParameter(,3,1,,id_target)

		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then							
			findTargetAssociations = true				
		end if		
		
		Set objRS = objCommand2.Execute()
		if not(objRS.EOF) then							
			findTargetAssociations = true				
		end if
		
		Set objRS = objCommand3.Execute()
		if not(objRS.EOF) then							
			findTargetAssociations = true				
		end if
		
		Set objRS = objCommand4.Execute()
		if not(objRS.EOF) then							
			findTargetAssociations = true				
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		Set objCommand4 = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function getListaTargetType()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaTargetType = null		
		strSQL = "SELECT * FROM target_type;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objTarget
			do while not objRS.EOF				
				strID = objRS("id")
				strDesc = objRS("descrizione")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop					
			Set getListaTargetType = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	
	Public Function renderTargetBox(resJsVar, idBoxSx, idBoxDx, labelSx, labelDx, targetType, arrSx, arrDx, bolCheckAutoT, bolAddDescTrans, objLang)
		on error resume next
	
'response.write("<br>targetType:"&targetType)
		resJsValues =""
		arrType = Split(Trim(targetType), ",", -1, 1)
		Set objDict = Server.CreateObject("Scripting.Dictionary")
		for each g in arrType
			objDict.add CInt(g), ""
'response.write("<br>g:"&g)
		next
	
		renderTargetBox=""		
		renderTargetBox=renderTargetBox&"<div style='float: left;margin-right: 10px;'>"
		renderTargetBox=renderTargetBox&"<span class='labelForm'>"&labelSx&"</span><br>"
		renderTargetBox=renderTargetBox&"<ul id='"&idBoxSx&"' style='list-style-type: none; margin: 0; float: left; margin-right: 10px; background: #fff; padding: 5px; width: 230px;height: 160px;border: 1px solid #727272;overflow: auto;'>"
		
		if (Instr(1, typename(arrSx), "dictionary", 1) > 0) then
			for each y in arrSx.Keys
'response.write("<br>arrSx(y).getTargetType():"&arrSx(y).getTargetType())
				if(objDict.exists(Cint(arrSx(y).getTargetType())))then
					ttype = ""
					bolOkAuto = true
					
					if(bolCheckAutoT) then
						if(Cint(arrSx(y).isAutomatic())=1) then
							bolOkAuto = false
						end if
					end if

'response.write("<br>bolOkAuto:"& bolOkAuto)						
					if(bolOkAuto) then
						Select Case arrSx(y).getTargetType()
						Case 1
							if(bolAddDescTrans) then
							ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_cat")&") "
							end if
							ttype = ttype&arrSx(y).getTargetDescrizione()
'response.write("<br>ttype:"&ttype&" -bolAddDescTrans:"& bolAddDescTrans)	
						Case 2
							if(bolAddDescTrans) then
							ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_prod")&") "
							end if
							ttype = ttype&arrSx(y).getTargetDescrizione()
						Case 3
							if(bolAddDescTrans) then
							ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_lang")&") "
							end if
							
							labelLang = Replace(arrSx(y).getTargetDescrizione(), "lang_", "", 1, -1, 1)
							imgLang="<img src='"&Application("baseroot") & "/editor/img/flag/flag-"&labelLang&".png"&"' border=0 hspace=2 vspace=0 align=top>"								
							if not(objLang.getTranslated("portal.header.label.desc_lang."&labelLang) = "") then
								labelLang = objLang.getTranslated("portal.header.label.desc_lang."&labelLang)
							end if
							ttype=ttype&imgLang&labelLang
						Case else
						end Select
						renderTargetBox=renderTargetBox&"<li class='ui-state-highlight' style='margin: 2px; padding: 2px; font-size: 11px; width: 200px; cursor:move;' id='"&y&"'>"&ttype&"</li>"
						resJsValues=resJsValues&y&"|"
					end if
				end if
			next
		end if		
		
		renderTargetBox=renderTargetBox&"</ul>"
		renderTargetBox=renderTargetBox&"</div>"
		renderTargetBox=renderTargetBox&"<div>"
		renderTargetBox=renderTargetBox&"<span class='labelForm'>"&labelDx&"</span><br>"
		renderTargetBox=renderTargetBox&"<ul id='"&idBoxDx&"' style='list-style-type: none; margin: 0; float:left; background: #fff; padding: 5px; width: 230px;height: 160px;border: 1px solid #727272;overflow: auto;'>"

'response.write("<br>renderTargetBox:"&renderTargetBox)	

		if (Instr(1, typename(arrDx), "dictionary", 1) > 0) AND not(isEmpty(arrDx)) then
			if (Instr(1, typename(arrSx), "dictionary", 1) > 0) then 
				for each x in arrDx.Keys
					if(objDict.exists(Cint(arrDx(x).getTargetType())))then
						ttype = ""
						bolOkAuto = true
						
						if(bolCheckAutoT) then
							if(Cint(arrDx(x).isAutomatic())=1) then
								bolOkAuto = false
							end if
						end if

						if(bolOkAuto) then				
							Select Case arrDx(x).getTargetType()
							Case 1
								if(bolAddDescTrans) then
								ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_cat")&") "
								end if
								ttype = ttype&arrDx(x).getTargetDescrizione()
							Case 2
								if(bolAddDescTrans) then
								ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_prod")&") "
								end if
								ttype = ttype&arrDx(x).getTargetDescrizione()
							Case 3
								if(bolAddDescTrans) then
								ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_lang")&") "
								end if
								
								labelLang = Replace(arrDx(x).getTargetDescrizione(), "lang_", "", 1, -1, 1)	
								imgLang="<img src='"&Application("baseroot") & "/editor/img/flag/flag-"&labelLang&".png"&"' border=0 hspace=2 vspace=0 align=top>"						
								if not(objLang.getTranslated("portal.header.label.desc_lang."&labelLang) = "") then
									labelLang = objLang.getTranslated("portal.header.label.desc_lang."&labelLang)
								end if
								ttype=ttype&imgLang&labelLang
							Case else
							end Select
					
							if(arrSx.Exists(x)) then
								bolExistTarget = true
							else
								bolExistTarget = false
							end if		
'response.write("<br>aaaa bolExistTarget:"& bolExistTarget)					
							if not(bolExistTarget) then
								renderTargetBox=renderTargetBox&"<li class='ui-state-default' style='margin: 2px; padding: 2px; font-size: 11px; width: 200px; cursor:move;' id='"&x&"'>"&ttype&"</li>"						
							end if 
						end if

'response.write("<br>aaaa ttype:"&ttype&" -bolAddDescTrans:"& bolAddDescTrans&" -bolOkAuto:"& bolOkAuto)

					end if
				next				
			else
				for each x in arrDx.Keys
					if(objDict.exists(Cint(arrDx(x).getTargetType())))then
						ttype = ""
						bolOkAuto = true
						
						if(bolCheckAutoT) then
							if(Cint(arrDx(x).isAutomatic())=1) then
								bolOkAuto = false
							end if
						end if

'response.write("<br>ttype:"&ttype&" -bolAddDescTrans:"& bolAddDescTrans&" -bolOkAuto:"& bolOkAuto)	


						if(bolOkAuto) then
							Select Case arrDx(x).getTargetType()
							Case 1
								if(bolAddDescTrans) then
								ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_cat")&") "
								end if
								ttype = ttype&arrDx(x).getTargetDescrizione()
							Case 2
								if(bolAddDescTrans) then
								ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_prod")&") "
								end if
								ttype = ttype&arrDx(x).getTargetDescrizione()
							Case 3
								if(bolAddDescTrans) then
								ttype = "("&objLang.getTranslated("backend.target.detail.table.select.option.target_type_lang")&") "
								end if
								
								labelLang = Replace(arrDx(x).getTargetDescrizione(), "lang_", "", 1, -1, 1)	
								imgLang="<img src='"&Application("baseroot") & "/editor/img/flag/flag-"&labelLang&".png"&"' border=0 hspace=2 vspace=0 align=top>"						
								if not(objLang.getTranslated("portal.header.label.desc_lang."&labelLang) = "") then
									labelLang = objLang.getTranslated("portal.header.label.desc_lang."&labelLang)
								end if
								ttype=ttype&imgLang&labelLang
							Case else
							end Select
							renderTargetBox=renderTargetBox&"<li class='ui-state-default' style='margin: 2px; padding: 2px; font-size: 11px; width: 200px; cursor:move;' id='"&x&"'>"&ttype&"</li>"
						end if
					end if
				next
			end if	
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=005")
		end if	
			
		renderTargetBox=renderTargetBox&"</ul>"
		renderTargetBox=renderTargetBox&"</div><br clear='both'/>"

		renderTargetBox=renderTargetBox&"<script>"
		renderTargetBox=renderTargetBox&"var "&resJsVar&" = '"&resJsValues&"';"
		
		renderTargetBox=renderTargetBox&"$(function () {"
			renderTargetBox=renderTargetBox&"$( '#"&idBoxSx&"').sortable({"
				renderTargetBox=renderTargetBox&"connectWith: ""ul#"&idBoxDx&""""
				renderTargetBox=renderTargetBox&",receive: function(event, ui) {"					
					renderTargetBox=renderTargetBox&resJsVar&"+=ui.item.attr('id')+'|';"	
				renderTargetBox=renderTargetBox&"}"
				renderTargetBox=renderTargetBox&",remove: function(event, ui) {"
					renderTargetBox=renderTargetBox&resJsVar&"="&resJsVar&".replace(ui.item.attr('id')+'|','');"
				renderTargetBox=renderTargetBox&"}"
			renderTargetBox=renderTargetBox&"}).disableSelection();"
			
			renderTargetBox=renderTargetBox&"$( '#"&idBoxDx&"').sortable({"
				renderTargetBox=renderTargetBox&"connectWith: ""ul#"&idBoxSx&""""
			renderTargetBox=renderTargetBox&"}).disableSelection();"
		renderTargetBox=renderTargetBox&"});"
		renderTargetBox=renderTargetBox&"</script>"
 
		if Err.number <> 0 then
			'response.write(Err.description)
		end if			
	End Function		
End Class
%>