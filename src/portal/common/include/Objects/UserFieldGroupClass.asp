<%
Class UserFieldGroupClass
	Private id
	Private description
	Private order
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strid)
		id = strid
	End Sub	
	
	Public Function getDescription()
		getDescription = description
	End Function
	
	Public Sub setDescription(strDescription)
		description = strDescription
	End Sub
	
	Public Function getOrder()
		getOrder = order
	End Function
	
	Public Sub setOrder(strOrder)
		order = strOrder
	End Sub
		
	Public Function getListUserFieldGroup()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListUserFieldGroup = null		
		strSQL = "SELECT * FROM user_fields_group ORDER BY `order`;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objUserFieldGroup
			do while not objRS.EOF				
				Set objUserFieldGroup = new UserFieldGroupClass
				strID = objRS("id")
				objUserFieldGroup.setID(strID)
				objUserFieldGroup.setDescription(objRS("description"))		
				objUserFieldGroup.setOrder(objRS("order"))		
				objDict.add strID, objUserFieldGroup
				objRS.moveNext()
			loop
			Set objUserFieldGroup = nothing							
			Set getListUserFieldGroup = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findUserFieldGroupById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findUserFieldGroupById = null		
		strSQL = "SELECT * FROM user_fields_group WHERE id=?;"
		
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
			Set objUserFieldGroup = new UserFieldGroupClass
			strID = objRS("id")
			objUserFieldGroup.setID(strID)
			objUserFieldGroup.setDescription(objRS("description"))		
			objUserFieldGroup.setOrder(objRS("order"))	
			Set findUserFieldGroupById = objUserFieldGroup
			Set objUserFieldGroup = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertUserFieldGroup(description, order, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
	
		strSQL = "INSERT INTO user_fields_group(description, `order`) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
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
		
	Public Sub modifyUserFieldGroup(id, description, order, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "UPDATE user_fields_group SET "
		strSQL = strSQL & "description=?,"
		strSQL = strSQL & "`order`=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
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
			
	Public Sub insertUserFieldGroupNoTransaction(description, order)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO user_fields_group(description, `order`) VALUES("
		strSQL = strSQL & "?,?);"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyUserFieldGroupNoTransaction(id, description, order)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		strSQL = "UPDATE user_fields_group SET "
		strSQL = strSQL & "description=?,"
		strSQL = strSQL & "`order`=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteUserFieldGroup(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM user_fields_group WHERE id=?;"

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
	
End Class
%>