<%
Class UserFieldClass
	Private id
	Private description
	Private idGroup
	Private objGroup
	Private order
	Private typeField
	Private typeContent
	Private maxLenght
	Private values
	Private required
	Private enabled
	Private useFor
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strid)
		id = strid
	End Sub	
	
	Public Function getDescription()
		getdescription = description
	End Function
	
	Public Sub setDescription(strDescription)
		description = strDescription
	End Sub
	
	Public Function getIdGroup()
		getIdGroup = idGroup
	End Function
	
	Public Sub setIdGroup(strIdGroup)
		idGroup = strIdGroup
	End Sub	
	
	Public Function getObjGroup()
		Set getObjGroup = objGroup
	End Function
	
	Public Sub setObjGroup(strObjGroup)
		Set objGroup = strObjGroup
	End Sub	
	
	Public Function getOrder()
		getOrder = order
	End Function
	
	Public Sub setOrder(strOrder)
		order = strOrder
	End Sub
	
	Public Function getTypeField()
		getTypeField = typeField
	End Function
	
	Public Sub setTypeField(strTypeField)
		typeField = strTypeField
	End Sub
	
	Public Function getTypeContent()
		getTypeContent = typeContent
	End Function
	
	Public Sub setTypeContent(strTypeContent)
		typeContent = strTypeContent
	End Sub	
	
	Public Function getMaxLenght()
		getMaxLenght = maxLenght
	End Function
	
	Public Sub setMaxLenght(strMaxLenght)
		maxLenght = strMaxLenght
	End Sub	
	
	Public Function getValues()
		getValues = values
	End Function
	
	Public Sub setValues(strValues)
		values = strValues
	End Sub
	
	Public Function getRequired()
		getRequired = required
	End Function
	
	Public Sub setRequired(bolRequired)
		required = bolRequired
	End Sub
	
	Public Function getEnabled()
		getEnabled = enabled
	End Function
	
	Public Sub setEnabled(bolEnabled)
		enabled = bolEnabled
	End Sub
	
	Public Function getUseFor()
		getUseFor = useFor
	End Function
	
	Public Sub setUseFor(intUseFor)
		useFor = intUseFor
	End Sub
		
	Public Function getListUserField(enabled, useFor)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListUserField = null		
		strSQL = "SELECT user_fields.*, user_fields_group.description as gdesc, user_fields_group.order as gorder FROM user_fields LEFT JOIN user_fields_group ON user_fields.id_group=user_fields_group.id"

		if not(isNull(enabled)) AND not(isNull(useFor)) then
			strSQL = strSQL & " WHERE"
			if not(isNull(enabled)) then
				strSQL = strSQL & " enabled=?"
			end if
	
			if not(isNull(useFor)) then
				arrUseFor = Split(useFor, ",", -1, 1)
				if(Ubound(arrUseFor) > 0) then
					strSQL = strSQL & " AND("
					for each e in arrUseFor
						strSQL = strSQL & " use_for=? OR"
					next
					strSQL = strSQL & ")"
					strSQL = Replace(strSQL, " OR)", ")", 1, -1, 1)
				end if
			end if
		end if
		
		strSQL = strSQL & " ORDER BY gorder, user_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if not(isNull(enabled)) then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		end if
		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				for each e in arrUseFor
					objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,e)
				next
			end if
		end if
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objUserField
			do while not objRS.EOF				
				Set objUserField = new UserFieldClass
				strID = objRS("id")
				objUserField.setID(strID)
				objUserField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objUserField.setIdGroup(strIdGroup)
				
				Set objGroup = new UserFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objUserField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objUserField.setOrder(objRS("order"))	
				objUserField.setTypeField(objRS("type"))
				objUserField.setTypeContent(objRS("type_content"))
				objUserField.setMaxLenght(objRS("max_lenght"))	
				objUserField.setValues(objRS("values"))	
				objUserField.setRequired(objRS("required"))	
				objUserField.setEnabled(objRS("enabled"))		
				objUserField.setUsefor(objRS("use_for"))	
				objDict.add strID, objUserField
				objRS.moveNext()
			loop
			Set objUserField = nothing							
			Set getListUserField = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findUserFieldById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findUserFieldById = null		
		strSQL = "SELECT user_fields.*, user_fields_group.description as gdesc, user_fields_group.order as gorder FROM user_fields LEFT JOIN user_fields_group ON user_fields.id_group=user_fields_group.id WHERE user_fields.id=?;"
		
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
			Set objUserField = new UserFieldClass
			strID = objRS("id")
			objUserField.setID(strID)
			objUserField.setDescription(objRS("description"))
			strIdGroup = objRS("id_group")
			objUserField.setIdGroup(strIdGroup)
			
			Set objGroup = new UserFieldGroupClass
			objGroup.setID(strIdGroup)
			objGroup.setDescription(objRS("gdesc"))
			objGroup.setOrder(objRS("gorder"))
			
			objUserField.setObjGroup(objGroup)
			objUserField.setOrder(objRS("order"))	
			objUserField.setTypeField(objRS("type"))
			objUserField.setTypeContent(objRS("type_content"))
			objUserField.setMaxLenght(objRS("max_lenght"))		
			objUserField.setValues(objRS("values"))	
			objUserField.setRequired(objRS("required"))	
			objUserField.setEnabled(objRS("enabled"))			
			objUserField.setUsefor(objRS("use_for"))	
			Set findUserFieldById = objUserField
			Set objUserField = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertUserField(description, idGroup, order, typeField, typeContent, values, required, enabled, maxLenght, useFor, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
	
		strSQL = "INSERT INTO user_fields(description, id_group, `type`, type_content, `values`, `order`, max_lenght, required, enabled, use_for) VALUES("
		strSQL = strSQL & "?,"

		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if

		strSQL = strSQL & "?,?,?,?,"

		if(isNull(maxLenght) OR maxLenght = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if	
		
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		if not isNull(idGroup) AND not(idGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,values)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,useFor)
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
		
	Public Sub modifyUserField(id, description, idGroup, order, typeField, typeContent, values, required, enabled, maxLenght, useFor, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "UPDATE user_fields SET "
		strSQL = strSQL & "description=?,"
		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "id_group=NULL,"
		else
			strSQL = strSQL & "id_group=?,"			
		end if
		strSQL = strSQL & "`type`=?,"
		strSQL = strSQL & "`type_content`=?,"
		strSQL = strSQL & "`values`=?,"
		strSQL = strSQL & "`order`=?,"	
		if(isNull(maxLenght) OR maxLenght = "") then
			strSQL = strSQL & "max_lenght=NULL,"
		else
			strSQL = strSQL & "max_lenght=?,"			
		end if
		strSQL = strSQL & "required=?,"		
		strSQL = strSQL & "enabled=?,"		
		strSQL = strSQL & "use_for=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		if not isNull(idGroup) AND not(idGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,values)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,useFor)
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
			
	Public Sub insertUserFieldNoTransaction(description, idGroup, order, typeField, typeContent, values, required, enabled, maxLenght, useFor)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO user_fields(description, id_group, `type`, type_content, `values`, `order`, max_lenght, required, enabled, use_for) VALUES("
		strSQL = strSQL & "?,"

		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if

		strSQL = strSQL & "?,?,?,?,"

		if(isNull(maxLenght) OR maxLenght = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if	
		
		strSQL = strSQL & "?,?,?);"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		if not isNull(idGroup) AND not(idGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,values)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,useFor)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyUserFieldNoTransaction(id, description, idGroup, order, typeField, typeContent, values, required, enabled, maxLenght, useFor)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		strSQL = "UPDATE user_fields SET "
		strSQL = strSQL & "description=?,"
		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "id_group=NULL,"
		else
			strSQL = strSQL & "id_group=?,"			
		end if
		strSQL = strSQL & "`type`=?,"
		strSQL = strSQL & "`type_content`=?,"
		strSQL = strSQL & "`values`=?,"
		strSQL = strSQL & "`order`=?,"	
		if(isNull(maxLenght) OR maxLenght = "") then
			strSQL = strSQL & "max_lenght=NULL,"
		else
			strSQL = strSQL & "max_lenght=?,"			
		end if
		strSQL = strSQL & "required=?,"		
		strSQL = strSQL & "enabled=?,"		
		strSQL = strSQL & "use_for=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		if not isNull(idGroup) AND not(idGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,values)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,useFor)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteUserField(id)
		on error resume next
		Dim objDB, strSQL, strSQL2, objRS, objConn		
		strSQL = "DELETE FROM user_fields WHERE id=?;" 
		strSQL2 = "DELETE FROM user_fields_match WHERE id_field=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		
		Dim objCommand, objCommand2
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQL2
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		
		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand2.Execute()
		end if	
		objCommand.Execute()	
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

	'************************* GESTIONE LISTA TYPE *******************************
		
	Public Function getListaTypeField()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaTypeField = null		
		strSQL = "SELECT * FROM user_fields_type;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("description")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaTypeField = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findTypeFieldById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findTypeFieldById = null		
		strSQL = "SELECT * FROM user_fields_type WHERE id=?;"
		
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
			findTypeFieldById = objRS("description")
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	'************************* GESTIONE LISTA TYPE CONTENT *******************************
		
	Public Function getListaTypeContent()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaTypeContent = null		
		strSQL = "SELECT * FROM user_fields_type_content;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("description")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaTypeContent = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findTypeContentById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findTypeContentById = null		
		strSQL = "SELECT * FROM user_fields_type_content WHERE id=?;"
		
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
			findTypeContentById = objRS("description")
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	'************************* GESTIONE FIELD MATCH *******************************
		
	Public Function findFieldMatch(idField, idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldMatch = null		
		strSQL = "SELECT * FROM user_fields_match WHERE id_field=? AND id_user=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			strID = objRS("id")
			strVal = objRS("value")		
			objDict.add "id", strID	
			objDict.add "value", strVal
			Set findFieldMatch = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findFieldMatchValue(idField, idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldMatchValue = null		
		strSQL = "SELECT * FROM user_fields_match WHERE id_field=? AND id_user=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			strVal = objRS("value")	
			findFieldMatchValue = strVal		
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findFieldMatchValueUnique(idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldMatchValueUnique = null		
		strSQL = "SELECT DISTINCT(value), automatic_user FROM user_fields_match LEFT JOIN utenti ON user_fields_match.id_user=utenti.id WHERE id_field=? AND automatic_user=0;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objDict = Server.CreateObject("Scripting.Dictionary")			
			do while not objRS.EOF
				strVal = objRS("value")
				if(Trim(strVal)<>"")then
				objDict.add strVal, ""
				end if
				objRS.moveNext()
			loop
							
			Set findFieldMatchValueUnique = objDict			
			Set objDict = nothing	
		end if

		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findFieldMatchAndDesc(idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldMatchAndDesc = null		
		strSQL = "SELECT value, user_fields.description as desc FROM user_fields_match LEFT JOIN user_fields ON user_fields_match.id_field=user_fields.id WHERE id_user=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idUser)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strDesc = objRS("desc")
				strVal = objRS("value")		
				objDict.add strDesc, strVal
				objRS.moveNext()
			loop
							
			Set findFieldMatchAndDesc = objDict			
			Set objDict = nothing	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertFieldMatch(idField, idUser, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO user_fields_match(id_field, id_user, value) VALUES("
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,value)
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
		
	Public Sub modifyFieldMatch(id, idField, idUser, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE user_fields_match SET "
		strSQL = strSQL & "value=?"
		strSQL = strSQL & " WHERE id_field=? AND id_user=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,value)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
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
		
	Public Sub deleteFieldMatch(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM user_fields_match WHERE id=?;"

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
		
	Public Sub deleteFieldMatchByUser(idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM user_fields_match WHERE id_user=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	'************************* GESTIONE FIELD RENDERING *******************************
	
	Public Function renderUserFieldHTML(userField,cssClass, idUser, defaultMatchValue, translator)	
		Dim fieldMatchValue, spitValues, keyPress, maxLenght, style
		
		fieldMatchValue = defaultMatchValue
		
		if not(idUser="") then
			on error resume next
				Set fieldMatchValue = findFieldMatch(userField.getID(),idUser)
				if (Instr(1, typename(fieldMatchValue), "dictionary", 1) > 0) then
					fieldMatchValue = fieldMatchValue.Item("value")
				end if
			if Err.number <> 0 then
				'response.write(Err.description)
			end if			
		end if
		
		on error resume next			
		keyPress = ""
		select Case userField.getTypeContent()
		Case 3		
			keyPress = " onkeypress=""javascript:return isInteger(event);"""
		Case 4		
			keyPress = " onkeypress=""javascript:return isDouble(event);"""
		Case Else
		End Select		
		
		maxLenght = ""		
		if not(userField.getMaxLenght()="") AND (userField.getMaxLenght()>0) then
			maxLenght = " maxlength="""&userField.getMaxLenght()&""""
		end if
		
		style = ""		
		if not(cssClass="") then
			style = " class="""&cssClass&""""
		end if

		
		renderUserFieldHTML = ""		

		select Case userField.getTypeField()
		Case 1
			renderUserFieldHTML = "<input type=""text"" name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" value="""&fieldMatchValue&""" "&style&" "&keyPress&maxLenght&" onfocus=""cleanInputField('"&getFieldPrefix()&userField.getID()&"');"" onBlur=""restoreInputField('"&getFieldPrefix()&userField.getID()&"','"&fieldMatchValue&"');""/>"

			if(CInt(userField.getTypeContent())=6) then
				renderUserFieldHTML = renderUserFieldHTML & "<script>"
				renderUserFieldHTML = renderUserFieldHTML & "$(function() {"
					renderUserFieldHTML = renderUserFieldHTML & "$('#"&getFieldPrefix()&userField.getID()&"').datepicker({"
						renderUserFieldHTML = renderUserFieldHTML & "dateFormat: 'dd/mm/yy',"
						renderUserFieldHTML = renderUserFieldHTML & "changeMonth: true,"
						renderUserFieldHTML = renderUserFieldHTML & "changeYear: true,"
						renderUserFieldHTML = renderUserFieldHTML & "yearRange: '1900:"&DatePart("yyyy",Date())&"'" 
					renderUserFieldHTML = renderUserFieldHTML & "});"
				renderUserFieldHTML = renderUserFieldHTML & "});"
				renderUserFieldHTML = renderUserFieldHTML & "</script>"
			end if
		Case 2
			renderUserFieldHTML = "<textarea name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
		Case 3
			renderUserFieldHTML = "<input type=""password"" name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" "&style&" value="""&fieldMatchValue&"""/>"
		Case 4			
			if(CInt(userField.getTypeContent())=5) then
				Dim key, objCountry 
				Set objCountry = New CountryClass

				renderUserFieldHTML = "<select name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" "&style&">"
				renderUserFieldHTML = renderUserFieldHTML & "<option value=""""></option>"
				if not(isNull(objCountry.findCountryListOnly("1,3")))then
					Set specialFieldValue = objCountry.findCountryListOnly("1,3")
					for each x in specialFieldValue
						key =  specialFieldValue(x).getCountryCode()
						selected = ""
						if (strComp(key, fieldMatchValue, 1) = 0) then selected=" selected" end if
						renderUserFieldHTML = renderUserFieldHTML & "<option value="""&key&""" "&selected&">"&translator.getTranslated("portal.commons.select.option.country."&key)&"</option>"     
					next
					Set specialFieldValue = nothing
				end if
				renderUserFieldHTML = renderUserFieldHTML & "</select>"
				Set objCountry = nothing
			else
				renderUserFieldHTML = "<select name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" "&style&">"
				renderUserFieldHTML = renderUserFieldHTML & "<option value=""""></option>"
				spitValues = Split(userField.getValues(),",")
				for each x in spitValues
					selected = ""
					if (strComp(Trim(x), fieldMatchValue, 1) = 0) then selected=" selected" end if
					label= Trim(x)
					if not(translator.getTranslated("portal.commons.user_field.label."&label)="") then label=translator.getTranslated("portal.commons.user_field.label."&label) end if
					renderUserFieldHTML = renderUserFieldHTML & "<OPTION VALUE="""&x&""" "&selected&">"&label&"</OPTION>"
				next
				renderUserFieldHTML = renderUserFieldHTML & "</select>"
			end if
		Case 5
			renderUserFieldHTML ="<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&userField.getID()&""">"
			renderUserFieldHTML = renderUserFieldHTML & "<select name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" multiple size="""&userField.getMaxLenght()&""" "&style&">"
			renderUserFieldHTML = renderUserFieldHTML & "<option value=""""></option>"			
			
			spitValues = Split(userField.getValues(),",")
			for each x in spitValues
				selected = ""
				'if (strComp(Trim(x), fieldMatchValue, 1) = 0) then selected=" selected" end if
				
				if not(fieldMatchValue = "") then
					spitMatchValues = Split(fieldMatchValue,",")
					for j=0 to Ubound(spitMatchValues)
						if(strComp(Trim(spitMatchValues(j)), Trim(x), 1) = 0) then
							selected=" selected"
							exit for
						end if
					next
				end if
				
				label= Trim(x)
				if not(translator.getTranslated("portal.commons.user_field.label."&label)="") then label=translator.getTranslated("portal.commons.user_field.label."&label) end if
				renderUserFieldHTML = renderUserFieldHTML & "<OPTION VALUE="""&x&""" "&selected&">"&label&"</OPTION>"
			next
			
			renderUserFieldHTML = renderUserFieldHTML & "</select>"
		Case 6			
			on error Resume Next
			if not(userField.getValues() = "") then
				spitValues = Split(userField.getValues(),",")
			end if
			K=1		
	
			'renderUserFieldHTML =renderUserFieldHTML & "<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&userField.getID()&""" id=""hidden_"&getFieldPrefix()&userField.getID()&"""/>"
			renderUserFieldHTML =renderUserFieldHTML & "<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&userField.getID()&""">"

			for each y in spitValues
				checked = ""
				if not(fieldMatchValue = "") then
					spitMatchValues = Split(fieldMatchValue,",")
					for j=0 to Ubound(spitMatchValues)
						if(strComp(Trim(spitMatchValues(j)), Trim(y), 1) = 0) then
							'checked = "checked"
							checked=" checked='checked'"
							exit for
						end if
					next
				end if
				newLine = ""
				if((k Mod 4) = 0) then newLine="<br/>" end if
				label= Trim(y)
				if not(translator.getTranslated("portal.commons.user_field.label."&label)="") then label=translator.getTranslated("portal.commons.user_field.label."&label) end if
				'renderUserFieldHTML =renderUserFieldHTML & label&"&nbsp;<input type=""checkbox"" "&style&" value="""&y&""" name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" "&checked&"/>"&newLine
				renderUserFieldHTML =renderUserFieldHTML & label&"&nbsp;<input type=""checkbox"" "&style&" value="""&y&""" name="""&getFieldPrefix()&userField.getID()&""" "&checked&"/>&nbsp;&nbsp;"&newLine
				k = k+1
			next
			
			if Err.number <> 0 then
				'response.write(Err.description)
			end if
		Case 7			
			on error Resume Next
			if not(userField.getValues() = "") then
				spitValues = Split(userField.getValues(),",")
			end if
			K=1
			
			'renderUserFieldHTML =renderUserFieldHTML & "<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&userField.getID()&""" id=""hidden_"&getFieldPrefix()&userField.getID()&"""/>"
			renderUserFieldHTML =renderUserFieldHTML & "<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&userField.getID()&""">"

			for each y in spitValues
				checked = ""
				if not(fieldMatchValue = "") then
					if (strComp(fieldMatchValue, Trim(y), 1) = 0) then checked=" checked='checked'" end if
				end if
				newLine = ""
				if((k Mod 4) = 0) then newLine="<br/>" end if
				label= Trim(y)
				if not(translator.getTranslated("portal.commons.user_field.label."&label)="") then label=translator.getTranslated("portal.commons.user_field.label."&label) end if
				'renderUserFieldHTML =renderUserFieldHTML & label&"&nbsp;<input type=""radio"" "&style&" value="""&y&""" name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" "&checked&"/>"&newLine
				renderUserFieldHTML =renderUserFieldHTML & label&"&nbsp;<input type=""radio"" "&style&" value="""&y&""" name="""&getFieldPrefix()&userField.getID()&""" "&checked&"/>&nbsp;&nbsp;"&newLine
				k = k+1
			next
			
			if Err.number <> 0 then
				'response.write(Err.description)
			end if
		Case 8
			renderUserFieldHTML = "<input type=""hidden"" value="""&fieldMatchValue&""" name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&"""/>"
		Case 9
			renderUserFieldHTML = "<input type=""file"" name="""&getFieldPrefix()&userField.getID()&""" id="""&getFieldPrefix()&userField.getID()&""" "&style&" />"
		Case Else
		End Select	
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function
	
	Public Function renderUserFieldJS(userField, formName, translator,defaultMatchValue, returnFalse)	
		on error resume next
		
		renderUserFieldJS = ""	

		renderReturnMode = "return;"

		if(returnFalse) then
			renderReturnMode = "return false;"
		end if		

		select Case userField.getTypeField()
		Case 1,2,3
			renderUserFieldJS = "var "&getFieldPrefix()&userField.getID()&"_default_match_values = """&defaultMatchValue&""";"
			renderUserFieldJS = renderUserFieldJS & "if(document."&formName&"."&getFieldPrefix()&userField.getID()&".value.toLowerCase() == "&getFieldPrefix()&userField.getID()&"_default_match_values.toLowerCase()){"
				renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".value = """";"
			renderUserFieldJS = renderUserFieldJS &"}"				
		
			if(CInt(userField.getRequired())=1)then
				renderUserFieldJS = renderUserFieldJS & "if(document."&formName&"."&getFieldPrefix()&userField.getID()&".value == """"){"
					renderUserFieldJS = renderUserFieldJS &"alert("""&translator.getTranslated("portal.commons.user_field.js.alert.insert_"&userField.getDescription())&""");"
					renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".focus();"
					renderUserFieldJS = renderUserFieldJS &renderReturnMode
				renderUserFieldJS = renderUserFieldJS &"}"			
			end if

			if(CInt(userField.getTypeContent())=2) then
				renderUserFieldJS = renderUserFieldJS &"if (document."&formName&"."&getFieldPrefix()&userField.getID()&".value.indexOf(""@"")<2 || document."&formName&"."&getFieldPrefix()&userField.getID()&".value.indexOf(""."")==-1 || document."&formName&"."&getFieldPrefix()&userField.getID()&".value.indexOf("" "")!=-1 || document."&formName&"."&getFieldPrefix()&userField.getID()&".value.length<6){"
					renderUserFieldJS = renderUserFieldJS &"alert("""&translator.getTranslated("portal.commons.user_field.js.alert.wrong_mail")&""");"
					renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".focus();"
					renderUserFieldJS = renderUserFieldJS &renderReturnMode
				renderUserFieldJS = renderUserFieldJS &"}"		
			end if

			if(CInt(userField.getTypeContent())=3) then
				renderUserFieldJS = renderUserFieldJS &"if(isNaN(document."&formName&"."&getFieldPrefix()&userField.getID()&".value)){"
					renderUserFieldJS = renderUserFieldJS &"alert("""&translator.getTranslated("portal.commons.user_field.js.alert.isnan_value")&""");"
					renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".focus();"
					renderUserFieldJS = renderUserFieldJS &renderReturnMode	
				renderUserFieldJS = renderUserFieldJS &"}"			
			end if

			if(CInt(userField.getTypeContent())=4) then
				renderUserFieldJS = renderUserFieldJS &"if(document."&formName&"."&getFieldPrefix()&userField.getID()&".value.length > 0 && (!checkDoubleFormatExt(document."&formName&"."&getFieldPrefix()&userField.getID()&".value) || document."&formName&"."&getFieldPrefix()&userField.getID()&".value.indexOf(""."")!=-1)){"
					renderUserFieldJS = renderUserFieldJS &"alert("""&translator.getTranslated("portal.commons.user_field.js.alert.isnan_value")&""");"
					renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".focus();"
					renderUserFieldJS = renderUserFieldJS &renderReturnMode
				renderUserFieldJS = renderUserFieldJS &"}"		
			end if

		Case 4
			if(CInt(userField.getRequired())=1)then
				renderUserFieldJS = "if(document."&formName&"."&getFieldPrefix()&userField.getID()&".options[document."&formName&"."&getFieldPrefix()&userField.getID()&".selectedIndex].value == """"){"
					renderUserFieldJS = renderUserFieldJS &"alert("""&translator.getTranslated("portal.commons.user_field.js.alert.insert_"&userField.getDescription())&""");"
					renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".focus();"
					renderUserFieldJS = renderUserFieldJS &renderReturnMode
				renderUserFieldJS = renderUserFieldJS &"}"		
			end if

		Case 5
			renderUserFieldJS = ""
			if(CInt(userField.getRequired())=1)then
				renderUserFieldJS = renderUserFieldJS & "var "&getFieldPrefix()&userField.getID()&"_hasselection = false;"				
				renderUserFieldJS = renderUserFieldJS & "for(k=0; k<document."&formName&"."&getFieldPrefix()&userField.getID()&".options.length; k++){"
					renderUserFieldJS = renderUserFieldJS & "if(document."&formName&"."&getFieldPrefix()&userField.getID()&".options[k].selected && document."&formName&"."&getFieldPrefix()&userField.getID()&".options[k].value != """"){"							
						renderUserFieldJS = renderUserFieldJS &getFieldPrefix()&userField.getID()&"_hasselection = true;"
						renderUserFieldJS = renderUserFieldJS &"break;"
					renderUserFieldJS = renderUserFieldJS & "}"					
				renderUserFieldJS = renderUserFieldJS & "}"
				renderUserFieldJS = renderUserFieldJS & "if(!"&getFieldPrefix()&userField.getID()&"_hasselection){"
					renderUserFieldJS = renderUserFieldJS &"alert("""&translator.getTranslated("portal.commons.user_field.js.alert.insert_"&userField.getDescription())&""");"
					renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".focus();"
					renderUserFieldJS = renderUserFieldJS &renderReturnMode
				renderUserFieldJS = renderUserFieldJS &"}"		
			end if

			renderUserFieldJS = renderUserFieldJS & "if(document."&formName&"."&getFieldPrefix()&userField.getID()&"){"
				renderUserFieldJS = renderUserFieldJS & "var "&getFieldPrefix()&userField.getID()&"_values = """";"
				renderUserFieldJS = renderUserFieldJS & "for(k=0; k<document."&formName&"."&getFieldPrefix()&userField.getID()&".options.length; k++){"
					renderUserFieldJS = renderUserFieldJS & "if(document."&formName&"."&getFieldPrefix()&userField.getID()&".options[k].selected){"							
						renderUserFieldJS = renderUserFieldJS &getFieldPrefix()&userField.getID()&"_values = "&getFieldPrefix()&userField.getID()&"_values + document."&formName&"."&getFieldPrefix()&userField.getID()&".options[k].value + "","";"
					renderUserFieldJS = renderUserFieldJS & "}"					
				renderUserFieldJS = renderUserFieldJS & "}"
				renderUserFieldJS = renderUserFieldJS &getFieldPrefix()&userField.getID()&"_values = "&getFieldPrefix()&userField.getID()&"_values.substring(0, "&getFieldPrefix()&userField.getID()&"_values.lastIndexOf(','));"
				renderUserFieldJS = renderUserFieldJS &"document."&formName&".hidden_"&getFieldPrefix()&userField.getID()&".value = "&getFieldPrefix()&userField.getID()&"_values;"	
				renderUserFieldJS = renderUserFieldJS & "for(i=document."&formName&"."&getFieldPrefix()&userField.getID()&".options.length-1; i>=0; i--){document."&formName&"."&getFieldPrefix()&userField.getID()&".remove(i);}"			
			renderUserFieldJS = renderUserFieldJS & "}"
		Case 6,7
			renderUserFieldJS = "var "&getFieldPrefix()&userField.getID()&"_values = """";"
			renderUserFieldJS = renderUserFieldJS &"if (document."&formName&"."&getFieldPrefix()&userField.getID()&"){"
				renderUserFieldJS = renderUserFieldJS &"if(document."&formName&"."&getFieldPrefix()&userField.getID()&" != null){"				
					renderUserFieldJS = renderUserFieldJS &"if(document."&formName&"."&getFieldPrefix()&userField.getID()&".length == null){"
						renderUserFieldJS = renderUserFieldJS &"if (document."&formName&"."&getFieldPrefix()&userField.getID()&".checked){"
							renderUserFieldJS = renderUserFieldJS &getFieldPrefix()&userField.getID()&"_values = "&getFieldPrefix()&userField.getID()&"_values + document."&formName&"."&getFieldPrefix()&userField.getID()&".value + "","";"
							renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".checked=false;"
							'renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&".disabled=true;"
						renderUserFieldJS = renderUserFieldJS &"}"						
					renderUserFieldJS = renderUserFieldJS &"}else{"
						renderUserFieldJS = renderUserFieldJS &"for (var i=0; i < document."&formName&"."&getFieldPrefix()&userField.getID()&".length; i++){"
							renderUserFieldJS = renderUserFieldJS &"if (document."&formName&"."&getFieldPrefix()&userField.getID()&"[i].checked){"
								renderUserFieldJS = renderUserFieldJS &getFieldPrefix()&userField.getID()&"_values = "&getFieldPrefix()&userField.getID()&"_values + document."&formName&"."&getFieldPrefix()&userField.getID()&"[i].value + "","";"
								renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&"[i].checked=false;"
								'renderUserFieldJS = renderUserFieldJS &"document."&formName&"."&getFieldPrefix()&userField.getID()&"[i].disabled=true;"
							renderUserFieldJS = renderUserFieldJS &"}"
						renderUserFieldJS = renderUserFieldJS &"}"						
					renderUserFieldJS = renderUserFieldJS &"}"
						
					renderUserFieldJS = renderUserFieldJS &getFieldPrefix()&userField.getID()&"_values = "&getFieldPrefix()&userField.getID()&"_values.substring(0, "&getFieldPrefix()&userField.getID()&"_values.lastIndexOf(','));"
				renderUserFieldJS = renderUserFieldJS &"}"
			renderUserFieldJS = renderUserFieldJS &"}"
			renderUserFieldJS = renderUserFieldJS &"document."&formName&".hidden_"&getFieldPrefix()&userField.getID()&".value = "&getFieldPrefix()&userField.getID()&"_values;"	
				
			if(CInt(userField.getRequired())=1)then
				renderUserFieldJS = renderUserFieldJS &"if ("&getFieldPrefix()&userField.getID()&"_values ==""""){"
				renderUserFieldJS = renderUserFieldJS &"alert("""&translator.getTranslated("portal.commons.user_field.js.alert.insert_"&userField.getDescription())&""");"
				renderUserFieldJS = renderUserFieldJS &renderReturnMode
				renderUserFieldJS = renderUserFieldJS &"}"
			end if	
		Case Else
		End Select
 
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function
	
	Public Function getFieldPrefix()
		getFieldPrefix = "userfield"
	End Function

	Function SortDictionary(objDict,intSort)
	  ' declare our variables
	  Dim dictKey, dictItem
	  Dim strDict()
	  Dim objKey
	  Dim strKey,strItem
	  Dim X,Y,Z
	  
	  'Set SortDictionary = null
	  
	  dictKey  = 1
	  dictItem = 2
	
	  ' get the dictionary count
	  Z = objDict.Count
	
	  ' we need more than one item to warrant sorting
	  If Z > 1 Then
		' create an array to store dictionary information
		ReDim strDict(Z,2)
		X = 0
		' populate the string array
		For Each objKey In objDict
			strDict(X,dictKey)  = CStr(objKey)
			strDict(X,dictItem) = CStr(objDict(objKey))
			X = X + 1
		Next
	
		' perform a a shell sort of the string array
		For X = 0 to (Z - 2)
		  For Y = X to (Z - 1)
			If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
				strKey  = strDict(X,dictKey)
				strItem = strDict(X,dictItem)
				strDict(X,dictKey)  = strDict(Y,dictKey)
				strDict(X,dictItem) = strDict(Y,dictItem)
				strDict(Y,dictKey)  = strKey
				strDict(Y,dictItem) = strItem
			End If
		  Next
		Next
	
		' erase the contents of the dictionary object
		objDict.RemoveAll
	
		' repopulate the dictionary with the sorted information
		For X = 0 to (Z - 1)
		  objDict.Add strDict(X,dictKey), strDict(X,dictItem)
		Next
	
	  End If
	  Set SortDictionary = objDict
	End Function
End Class
%>