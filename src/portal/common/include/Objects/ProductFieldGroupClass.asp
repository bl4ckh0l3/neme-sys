<%
Class ProductFieldGroupClass
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
		
	Public Function getListProductFieldGroup()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListProductFieldGroup = null		
		strSQL = "SELECT * FROM product_fields_group ORDER BY `order`;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objProductFieldGroup
			do while not objRS.EOF				
				Set objProductFieldGroup = new ProductFieldGroupClass
				strID = objRS("id")
				objProductFieldGroup.setID(strID)
				objProductFieldGroup.setDescription(objRS("description"))		
				objProductFieldGroup.setOrder(objRS("order"))		
				objDict.add strID, objProductFieldGroup
				objRS.moveNext()
			loop
			Set objProductFieldGroup = nothing							
			Set getListProductFieldGroup = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findProductFieldGroupById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findProductFieldGroupById = null		
		strSQL = "SELECT * FROM product_fields_group WHERE id=?;"
		
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
			Set objProductFieldGroup = new ProductFieldGroupClass
			strID = objRS("id")
			objProductFieldGroup.setID(strID)
			objProductFieldGroup.setDescription(objRS("description"))		
			objProductFieldGroup.setOrder(objRS("order"))	
			Set findProductFieldGroupById = objProductFieldGroup
			Set objProductFieldGroup = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertProductFieldGroup(description, order, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
	
		strSQL = "INSERT INTO product_fields_group(description, `order`) VALUES("
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
		
	Public Sub modifyProductFieldGroup(id, description, order, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "UPDATE product_fields_group SET "
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
			
	Public Sub insertProductFieldGroupNoTransaction(description, order)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO product_fields_group(description, `order`) VALUES("
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
		
	Public Sub modifyProductFieldGroupNoTransaction(id, description, order)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		strSQL = "UPDATE product_fields_group SET "
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
		
	Public Sub deleteProductFieldGroup(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_group WHERE id=?;"

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