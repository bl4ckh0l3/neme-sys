<%
Class ProductFieldClass
	Private id
	Private description
	Private idGroup
	Private objGroup
	Private order
	Private typeField
	Private typeContent
	Private maxLenght
	Private required
	Private enabled
	Private editable
	'*** field aggiuntivi per gestione ordine
	Private foCounter
	Private idOrder
	Private idProd
	Private qtaProd
	Private selValue
	'*** field aggiuntivi per gestione carrello
	Private fcCounter
	Private idCard
	
	
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
	
	Public Function getEditable()
		getEditable = editable
	End Function
	
	Public Sub setEditable(bolEditable)
		editable = bolEditable
	End Sub
	
	'*********************************************************
	
	Public Function getFoCounter()
		getFoCounter = foCounter
	End Function
	
	Public Sub setFoCounter(strFoCounter)
		foCounter = strFoCounter
	End Sub
	
	Public Function getIdOrder()
		getIdOrder = idOrder
	End Function
	
	Public Sub setIdOrder(strIdOrder)
		idOrder = strIdOrder
	End Sub
	
	Public Function getIdProd()
		getIdProd = idProd
	End Function
	
	Public Sub setIdProd(strIdProd)
		idProd = strIdProd
	End Sub
	
	Public Function getQtaProd()
		getQtaProd = qtaProd
	End Function
	
	Public Sub setQtaProd(strQtaProd)
		qtaProd = strQtaProd
	End Sub
	
	Public Function getSelValue()
		getSelValue = selValue
	End Function
	
	Public Sub setSelValue(strSelValue)
		selValue = strSelValue
	End Sub
	
	'*********************************************************
	
	Public Function getFcCounter()
		getFcCounter = fcCounter
	End Function
	
	Public Sub setFcCounter(strFcCounter)
		fcCounter = strFcCounter
	End Sub
	
	Public Function getIdCard()
		getIdCard = idCard
	End Function
	
	Public Sub setIdCard(strIdCard)
		idCard = strIdCard
	End Sub	
	
	
	'************************* GESTIONE PRODUCT FIELDS *******************************	
		
	Public Function getListProductField(enabled)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListProductField = null		
		strSQL = "SELECT product_fields.*, product_fields_group.description as gdesc, product_fields_group.order as gorder FROM product_fields LEFT JOIN product_fields_group ON product_fields.id_group=product_fields_group.id"

		if not(isNull(enabled)) then
			strSQL = strSQL & " WHERE enabled=?"
		end if
		
		strSQL = strSQL & " ORDER BY gorder, product_fields.order;"
		
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
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objProductField
			do while not objRS.EOF				
				Set objProductField = new ProductFieldClass
				strID = objRS("id")
				objProductField.setID(strID)
				objProductField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objProductField.setIdGroup(strIdGroup)
				
				Set objGroup = new ProductFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objProductField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objProductField.setOrder(objRS("order"))	
				objProductField.setTypeField(objRS("type"))
				objProductField.setTypeContent(objRS("type_content"))
				objProductField.setMaxLenght(objRS("max_lenght"))	
				objProductField.setRequired(objRS("required"))	
				objProductField.setEnabled(objRS("enabled"))		
				objProductField.setEditable(objRS("editable"))	
				objDict.add strID, objProductField
				objRS.moveNext()
			loop
			Set objProductField = nothing							
			Set getListProductField = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function getListProductField4Prod(idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListProductField4Prod = null		
		strSQL = "SELECT product_fields.*, product_fields_group.description as gdesc, product_fields_group.order as gorder, product_fields_match.id_prod, product_fields_match.value FROM product_fields"
		strSQL = strSQL & " LEFT JOIN product_fields_group ON product_fields.id_group=product_fields_group.id"
		strSQL = strSQL & " LEFT JOIN product_fields_match ON product_fields.id = product_fields_match.id_field AND product_fields_match.id_prod=?"
		strSQL = strSQL & " WHERE enabled=1"			
		strSQL = strSQL & " ORDER BY gorder, product_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objProductField
			do while not objRS.EOF				
				Set objProductField = new ProductFieldClass
				strID = objRS("id")
				objProductField.setID(strID)
				objProductField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objProductField.setIdGroup(strIdGroup)
				
				Set objGroup = new ProductFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objProductField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objProductField.setOrder(objRS("order"))	
				objProductField.setTypeField(objRS("type"))
				objProductField.setTypeContent(objRS("type_content"))
				objProductField.setMaxLenght(objRS("max_lenght"))	
				objProductField.setRequired(objRS("required"))	
				objProductField.setEnabled(objRS("enabled"))
				objProductField.setEditable(objRS("editable"))	
				objProductField.setIdProd(objRS("id_prod"))
				objProductField.setSelValue(objRS("value"))
				objDict.add strID, objProductField
				objRS.moveNext()
			loop
			
			Set objProductField = nothing							
			Set getListProductField4Prod = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListProductField4ProdActive(idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListProductField4ProdActive = null		
		strSQL = "SELECT product_fields.*, product_fields_group.description as gdesc, product_fields_group.order as gorder, product_fields_match.id_prod, product_fields_match.value FROM product_fields"
		strSQL = strSQL & " LEFT JOIN product_fields_group ON product_fields.id_group=product_fields_group.id"
		strSQL = strSQL & " LEFT JOIN product_fields_match ON product_fields.id = product_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND product_fields_match.id_prod=?"			
		strSQL = strSQL & " ORDER BY gorder, product_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objProductField
			do while not objRS.EOF				
				Set objProductField = new ProductFieldClass
				strID = objRS("id")
				objProductField.setID(strID)
				objProductField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objProductField.setIdGroup(strIdGroup)
				
				Set objGroup = new ProductFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objProductField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objProductField.setOrder(objRS("order"))	
				objProductField.setTypeField(objRS("type"))
				objProductField.setTypeContent(objRS("type_content"))
				objProductField.setMaxLenght(objRS("max_lenght"))		
				objProductField.setRequired(objRS("required"))	
				objProductField.setEnabled(objRS("enabled"))
				objProductField.setEditable(objRS("editable"))	
				objProductField.setIdProd(objRS("id_prod"))
				objProductField.setSelValue(objRS("value"))
				objDict.add strID, objProductField
				objRS.moveNext()
			loop
			
			Set objProductField = nothing							
			Set getListProductField4ProdActive = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListProductField4ProdActiveByType(idProd, typeP)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListProductField4ProdActiveByType = null		
		strSQL = "SELECT product_fields.*, product_fields_group.description as gdesc, product_fields_group.order as gorder, product_fields_match.id_prod, product_fields_match.value FROM product_fields"
		strSQL = strSQL & " LEFT JOIN product_fields_group ON product_fields.id_group=product_fields_group.id"
		strSQL = strSQL & " LEFT JOIN product_fields_match ON product_fields.id = product_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND product_fields_match.id_prod=?"	
		if not(isNull(typeP)) then
			strSQL = strSQL & " AND type IN("&typeP&")"	
		end if		
		strSQL = strSQL & " ORDER BY gorder, product_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objProductField
			do while not objRS.EOF				
				Set objProductField = new ProductFieldClass
				strID = objRS("id")
				objProductField.setID(strID)
				objProductField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objProductField.setIdGroup(strIdGroup)
				
				Set objGroup = new ProductFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objProductField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objProductField.setOrder(objRS("order"))	
				objProductField.setTypeField(objRS("type"))
				objProductField.setTypeContent(objRS("type_content"))
				objProductField.setMaxLenght(objRS("max_lenght"))		
				objProductField.setRequired(objRS("required"))	
				objProductField.setEnabled(objRS("enabled"))
				objProductField.setEditable(objRS("editable"))	
				objProductField.setIdProd(objRS("id_prod"))
				objProductField.setSelValue(objRS("value"))
				objDict.add strID, objProductField
				objRS.moveNext()
			loop
			
			Set objProductField = nothing							
			Set getListProductField4ProdActiveByType = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function hasListProductField4ProdActive(idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		hasListProductField4ProdActive = 0		
		strSQL = "SELECT count(*) as id FROM product_fields"
		strSQL = strSQL & " LEFT JOIN product_fields_group ON product_fields.id_group=product_fields_group.id"
		strSQL = strSQL & " LEFT JOIN product_fields_match ON product_fields.id = product_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND product_fields_match.id_prod=?"	
		strSQL = strSQL & " AND (product_fields.type not in(1,2,7,8,9) OR (product_fields.type in(1,2,8,9) AND product_fields.editable=1));"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			hasListProductField4ProdActive = objRS("id")			
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findProductFieldById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findProductFieldById = null		
		strSQL = "SELECT product_fields.*, product_fields_group.description as gdesc, product_fields_group.order as gorder FROM product_fields LEFT JOIN product_fields_group ON product_fields.id_group=product_fields_group.id WHERE product_fields.id=?;"
		
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
			Set objProductField = new ProductFieldClass
			strID = objRS("id")
			objProductField.setID(strID)
			objProductField.setDescription(objRS("description"))
			strIdGroup = objRS("id_group")
			objProductField.setIdGroup(strIdGroup)
			
			Set objGroup = new ProductFieldGroupClass
			objGroup.setID(strIdGroup)
			objGroup.setDescription(objRS("gdesc"))
			objGroup.setOrder(objRS("gorder"))
			
			objProductField.setObjGroup(objGroup)
			objProductField.setOrder(objRS("order"))	
			objProductField.setTypeField(objRS("type"))
			objProductField.setTypeContent(objRS("type_content"))
			objProductField.setMaxLenght(objRS("max_lenght"))		
			objProductField.setRequired(objRS("required"))	
			objProductField.setEnabled(objRS("enabled"))
			objProductField.setEditable(objRS("editable"))		
			Set findProductFieldById = objProductField
			Set objProductField = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Function insertProductField(description, idGroup, order, typeField, typeContent, required, enabled, editable, maxLenght, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		insertProductField = -1
	
		strSQL = "INSERT INTO product_fields(description, id_group, `type`, type_content, `order`, max_lenght, required, enabled, editable) VALUES("
		strSQL = strSQL & "?,"

		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if

		strSQL = strSQL & "?,?,?,"

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
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,editable)
		objCommand.Execute
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(product_fields.id) as id FROM product_fields")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertProductField = objRS("id")	
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
		
	Public Sub modifyProductField(id, description, idGroup, order, typeField, typeContent, required, enabled, editable, maxLenght, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "UPDATE product_fields SET "
		strSQL = strSQL & "description=?,"
		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "id_group=NULL,"
		else
			strSQL = strSQL & "id_group=?,"			
		end if
		strSQL = strSQL & "`type`=?,"
		strSQL = strSQL & "`type_content`=?,"
		strSQL = strSQL & "`order`=?,"	
		if(isNull(maxLenght) OR maxLenght = "") then
			strSQL = strSQL & "max_lenght=NULL,"
		else
			strSQL = strSQL & "max_lenght=?,"			
		end if
		strSQL = strSQL & "required=?,"		
		strSQL = strSQL & "enabled=?,"
		strSQL = strSQL & "editable=?"
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
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,editable)
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
			
	Public Function insertProductFieldNoTransaction(description, idGroup, order, typeField, typeContent, required, enabled, editable, maxLenght)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		insertProductFieldNoTransaction = -1
		
		strSQL = "INSERT INTO product_fields(description, id_group, `type`, type_content, `order`, max_lenght, required, enabled, editable) VALUES("
		strSQL = strSQL & "?,"

		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if

		strSQL = strSQL & "?,?,?,"

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
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,editable)
		objCommand.Execute()		
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(product_fields.id) as id FROM product_fields")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertProductFieldNoTransaction = objRS("id")	
		end if		
		Set objRS = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyProductFieldNoTransaction(id, description, idGroup, order, typeField, typeContent, required, enabled, editable, maxLenght)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		strSQL = "UPDATE product_fields SET "
		strSQL = strSQL & "description=?,"
		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "id_group=NULL,"
		else
			strSQL = strSQL & "id_group=?,"			
		end if
		strSQL = strSQL & "`type`=?,"
		strSQL = strSQL & "`type_content`=?,"
		strSQL = strSQL & "`order`=?,"	
		if(isNull(maxLenght) OR maxLenght = "") then
			strSQL = strSQL & "max_lenght=NULL,"
		else
			strSQL = strSQL & "max_lenght=?,"			
		end if
		strSQL = strSQL & "required=?,"	
		strSQL = strSQL & "enabled=?,"
		strSQL = strSQL & "editable=?"
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
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,editable)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteProductField(id)
		on error resume next
		Dim objDB, strSQL, strSQL2, objRS, objConn		
		strSQL = "DELETE FROM product_fields WHERE id=?;" 
		strSQL2 = "DELETE FROM product_fields_match WHERE id_field=?;"
		strSQL3 = "DELETE FROM product_fields_values WHERE id_field=?;"

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
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQL2
		objCommand3.CommandText = strSQL3
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)
		
		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand2.Execute()
			objCommand3.Execute()
		end if	
		objCommand.Execute()	

		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		
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

	
	'************************* GESTIONE PRODUCT FIELDS VALUES *******************************

	Public Function getListProductFieldValues(idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListProductFieldValues = null		
		strSQL = "SELECT * FROM product_fields_values "
		strSQL = strSQL & " WHERE id_field=?"		
		strSQL = strSQL & " ORDER BY `order`;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objProductField
			do while not objRS.EOF				
				strID = objRS("id_field")
				strValue = objRS("value")		
				objDict.add strValue, strID
				objRS.moveNext()
			loop						
			Set getListProductFieldValues = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Sub insertProductFieldValue(idField, strValue, iOrder, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO product_fields_values(id_field, `value`,`order`) VALUES("
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,iOrder)
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
		
	Public Sub modifyProductFieldValue(idField, strValue, iOrder, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE product_fields_values SET "
		strSQL = strSQL & "`value`=?,"
		strSQL = strSQL & "`order`=?"
		strSQL = strSQL & " WHERE id_field=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,iOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
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
		
	Public Sub deleteProductFieldValue(idField, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_values WHERE id_field=? AND `value`=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
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
		
	Public Sub deleteProductFieldValueNoTransaction(idField, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_values WHERE id_field=? AND `value`=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteProductFieldValueByField(idField, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_values WHERE id_field=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
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
		
	Public Sub deleteProductFieldValueByFieldNoTransaction(idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_values WHERE id_field=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Execute()
		Set objCommand = Nothing
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
		strSQL = "SELECT * FROM product_fields_type;"
		
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
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findTypeFieldById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findTypeFieldById = null		
		strSQL = "SELECT * FROM product_fields_type WHERE id=?;"
		
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
		strSQL = "SELECT * FROM product_fields_type_content;"
		
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
		strSQL = "SELECT * FROM product_fields_type_content WHERE id=?;"
		
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
		
	Public Function findFieldMatch(idField, idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldMatch = null		
		strSQL = "SELECT * FROM product_fields_match WHERE id_field=? AND id_prod=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
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
		
	Public Function findFieldMatchValue(idField, idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldMatchValue = null		
		strSQL = "SELECT * FROM product_fields_match WHERE id_field=? AND id_prod=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
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
	
	Public Sub insertFieldMatch(idField, idProd, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO product_fields_match(id_field, id_prod, `value`) VALUES("
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
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
		
	Public Sub modifyFieldMatch(idField, idProd, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE product_fields_match SET "
		strSQL = strSQL & "`value`=?"
		strSQL = strSQL & " WHERE id_field=? AND id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
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
		
	Public Sub deleteFieldMatch(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_match WHERE id=?;"

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
		
	Public Sub deleteFieldMatchNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_match WHERE id=?;"

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
		
	Public Sub deleteFieldMatchByProd(idProd, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_match WHERE id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
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
		
	Public Sub deleteFieldMatchByProdNoTransaction(idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_match WHERE id_prod=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	
	'************************* GESTIONE FIELD VALUE MATCH *******************************
	
	Public Function findFieldValueMatch(idField, idProd, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldValueMatch = null		
		strSQL = "SELECT * FROM product_fields_value_match WHERE id_field=? AND id_prod=? AND `value`=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then	
			iQtaProd = objRS("qta_prod")	
			findFieldValueMatch = iQtaProd
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findListFieldValueMatch(idField, idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findListFieldValueMatch = null		
		strSQL = "SELECT * FROM product_fields_value_match WHERE id_field=? AND id_prod=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then					
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				value = objRS("value")
				iQtaProd = objRS("qta_prod")
				objDict.add value, iQtaProd
				objRS.moveNext()
			loop
							
			Set findListFieldValueMatch = objDict			
			Set objDict = nothing	
			
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertFieldValueMatch(idField, idProd, iQtaProd, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO product_fields_value_match(id_field, id_prod, `qta_prod`, `value`) VALUES("
		strSQL = strSQL & "?,?,"

		if(isNull(iQtaProd) OR iQtaProd = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if	
		strSQL = strSQL & "?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		if not isNull(iQtaProd) AND not(iQtaProd = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,iQtaProd)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
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
		
	Public Sub modifyFieldValueMatch(idField, idProd, iQtaProd, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE product_fields_value_match SET "
		if(isNull(iQtaProd) OR iQtaProd = "") then
			strSQL = strSQL & "`qta_prod`=NULL"
		else
			strSQL = strSQL & "`qta_prod`=?"			
		end if
		strSQL = strSQL & " WHERE id_field=? AND id_prod=? AND `value`=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if not isNull(iQtaProd) AND not(iQtaProd = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,iQtaProd)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
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
		
	Public Sub modifyFieldValueMatchNoTransaction(idField, idProd, iQtaProd, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE product_fields_value_match SET "
		if(isNull(iQtaProd) OR iQtaProd = "") then
			strSQL = strSQL & "`qta_prod`=NULL"
		else
			strSQL = strSQL & "`qta_prod`=?"			
		end if
		strSQL = strSQL & " WHERE id_field=? AND id_prod=? AND `value`=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if not isNull(iQtaProd) AND not(iQtaProd = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,iQtaProd)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Execute()
		Set objCommand = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		


	Public Function changeQtaFieldValueMatchNoTransaction(idField, idProd, strValue, qta, oldQta)
		on error resume next
		changeQtaFieldValueMatchNoTransaction = -1
		
		Dim objDB, strSQL, objRS, newQta
		Dim objConn
		
		newQta = CInt(oldQta) - CInt(qta)			
		
		strSQL = "UPDATE product_fields_value_match SET "
		strSQL = strSQL & "`qta_prod`=?"	
		strSQL = strSQL & " WHERE id_field=? AND id_prod=? AND `value`=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,newQta)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Execute()	
		
		if(newQta = 0)then
			changeQtaFieldValueMatchNoTransaction = 0
		elseif(newQta > 0)then
			changeQtaFieldValueMatchNoTransaction = 1
		end if
						
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Function changeQtaFieldValueMatch(idField, idProd, strValue, qta, oldQta, objConn)
		on error resume next
		changeQtaFieldValueMatch = -1
		
		Dim objDB, strSQL, objRS, newQta		
		newQta = CInt(oldQta) - CInt(qta)			
		
		strSQL = "UPDATE product_fields_value_match SET "
		strSQL = strSQL & "`qta_prod`=?"
		strSQL = strSQL & " WHERE id_field=? AND id_prod=? AND `value`=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,newQta)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Execute()
		Set objCommand = Nothing
		
		if(newQta = 0)then
			changeQtaFieldValueMatch = 0
		elseif(newQta > 0)then
			changeQtaFieldValueMatch = 1
		end if
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	
	Public Sub deleteFieldValueMatch(idField, idProd, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_value_match WHERE id_field=? AND id_prod=? AND `value`=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
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
		
	Public Sub deleteFieldValueMatchNoTransaction(idField, idProd, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_value_match WHERE id_field=? AND id_prod=? AND `value`=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteFieldValueMatchByProd(idProd, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_value_match WHERE id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
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
		
	Public Sub deleteFieldValueMatchByProdNoTransaction(idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_value_match WHERE id_prod=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	
	'************************* GESTIONE FIELD CORRELATI VALUE MATCH *******************************
	
	Public Function findFieldRelValueMatch(idProd, idField, strValue, idFieldRel, strValueRel)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldRelValueMatch = null		
		strSQL = "SELECT * FROM product_fields_rel_value_match WHERE id_prod=? AND id_field=? AND `field_val`=? AND id_field_rel=? AND `field_rel_val`=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idFieldRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValueRel)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then	
			iQtaProd = objRS("qta_rel")	
			findFieldRelValueMatch = iQtaProd
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findListFieldRelValueMatch(idProd, idField, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findListFieldRelValueMatch = null		
		strSQL = "SELECT product_fields_rel_value_match.*, product_fields.description AS field_rel_desc FROM product_fields_rel_value_match"
		strSQL = strSQL & " LEFT JOIN product_fields ON product_fields_rel_value_match.id_field_rel = product_fields.id"
		strSQL = strSQL & " WHERE id_prod=? AND id_field=? AND `field_val`=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then					
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objValue = Server.CreateObject("Scripting.Dictionary")
				id_prod = objRS("id_prod")
				id_field = objRS("id_field")
				field_val = objRS("field_val")
				id_field_rel = objRS("id_field_rel")
				field_rel_val = objRS("field_rel_val")
				field_rel_desc = objRS("field_rel_desc")
				iQtaProd = objRS("qta_rel")
				value = id_prod&"|"&id_field&"|"&field_val&"|"&id_field_rel&"|"&field_rel_val
				objValue.add "id_prod", id_prod
				objValue.add "id_field", id_field
				objValue.add "field_val", field_val
				objValue.add "id_field_rel", id_field_rel
				objValue.add "field_rel_val", field_rel_val
				objValue.add "field_rel_desc", field_rel_desc
				objValue.add "qta_rel", iQtaProd
				objDict.add value, objValue
				objRS.moveNext()
			loop
							
			Set findListFieldRelValueMatch = objDict			
			Set objDict = nothing	
			
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertFieldRelValueMatch(idProd, idField, strValue, idFieldRel, strValueRel, iQtaProd, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO product_fields_rel_value_match(id_prod, id_field, `field_val`, id_field_rel, `field_rel_val`, `qta_rel`) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFieldRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValueRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,iQtaProd)
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
		
	Public Sub modifyFieldRelValueMatch(idProd, idField, strValue, idFieldRel, strValueRel, iQtaProd, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE product_fields_rel_value_match SET "
		strSQL = strSQL & "`qta_rel`=?"			
		strSQL = strSQL & " WHERE id_prod=? AND id_field=? AND `field_val`=? AND `id_field_rel`=? AND `field_rel_val`=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,iQtaProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFieldRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValueRel)
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
		
	Public Sub modifyFieldRelValueMatchNoTransaction(idProd, idField, strValue, idFieldRel, strValueRel, iQtaProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE product_fields_rel_value_match SET "
		strSQL = strSQL & "`qta_rel`=?"			
		strSQL = strSQL & " WHERE id_prod=? AND id_field=? AND `field_val`=? AND `id_field_rel`=? AND `field_rel_val`=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,iQtaProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFieldRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValueRel)
		objCommand.Execute()
		Set objCommand = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		


	Public Function changeQtaFieldRelValueMatchNoTransaction(idProd, idField, strValue, idFieldRel, strValueRel, qta, oldQta)
		on error resume next
		changeQtaFieldRelValueMatchNoTransaction = -1
		
		Dim objDB, strSQL, objRS, newQta
		Dim objConn
		
		newQta = CInt(oldQta) - CInt(qta)			
		
		strSQL = "UPDATE product_fields_rel_value_match SET "
		strSQL = strSQL & "`qta_rel`=?"	
		strSQL = strSQL & " WHERE id_prod=? AND id_field=? AND `field_val`=? AND `id_field_rel`=? AND `field_rel_val`=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,newQta)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFieldRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValueRel)
		objCommand.Execute()	
		
		if(newQta = 0)then
			changeQtaFieldRelValueMatchNoTransaction = 0
		elseif(newQta > 0)then
			changeQtaFieldRelValueMatchNoTransaction = 1
		end if
						
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Function changeQtaFieldRelValueMatch(idProd, idField, strValue, idFieldRel, strValueRel, qta, oldQta, objConn)
		on error resume next
		changeQtaFieldRelValueMatch = -1

		Dim objDB, strSQL, objRS, newQta		
		newQta = CInt(oldQta) - CInt(qta)	
		
		strSQL = "UPDATE product_fields_rel_value_match SET "
		strSQL = strSQL & "`qta_rel`=?"	
		strSQL = strSQL & " WHERE id_prod=? AND id_field=? AND `field_val`=? AND `id_field_rel`=? AND `field_rel_val`=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,newQta)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFieldRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValueRel)
		objCommand.Execute()
		Set objCommand = Nothing
		
		if(newQta = 0)then
			changeQtaFieldRelValueMatch = 0
		elseif(newQta > 0)then
			changeQtaFieldRelValueMatch = 1
		end if
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	
	Public Sub deleteFieldRelValueMatch(idProd, idField, strValue, idFieldRel, strValueRel, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_rel_value_match WHERE id_prod=? AND id_field=? AND `field_val`=? AND `id_field_rel`=? AND `field_rel_val`=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFieldRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValueRel)
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
		
	Public Sub deleteFieldRelValueMatchNoTransaction(idProd, idField, strValue, idFieldRel, strValueRel)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_rel_value_match WHERE id_prod=? AND id_field=? AND `field_val`=? AND `id_field_rel`=? AND `field_rel_val`=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFieldRel)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValueRel)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteFieldRelValueMatchByProd(idProd, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_rel_value_match WHERE id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
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
		
	Public Sub deleteFieldRelValueMatchByProdNoTransaction(idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_rel_value_match WHERE id_prod=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		

	'************************* GESTIONE FIELD X ORDER *******************************
		
	Public Function findFieldXOrder(counter, idOrder, idProd, idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldXOrder = null		
		strSQL = "SELECT * FROM product_fields_x_order WHERE `counter`=? AND id_order =? AND id_prod=? AND id_field=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,counter)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objProductField = new ProductFieldClass
			strID = objRS("id_field")
			objProductField.setID(strID)
			objProductField.setFoCounter(objRS("counter"))	
			objProductField.setIdOrder(objRS("id_order"))			
			objProductField.setIdProd(objRS("id_prod"))	
			objProductField.setQtaProd(objRS("qta_prod"))
			objProductField.setSelValue(objRS("value"))
			Set findFieldXOrder = objProductField			
			Set objProductField = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findListFieldXOrderByProd(counter, idOrder, idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findListFieldXOrderByProd = null		
		strSQL = "SELECT product_fields_x_order.*, product_fields.`type`, product_fields.description as `desc` FROM product_fields_x_order LEFT JOIN product_fields ON product_fields_x_order.id_field=product_fields.id"
		strSQL = strSQL & " WHERE id_order =? AND id_prod=?"
		if not(isNull(counter))then
			strSQL = strSQL & " AND counter=?"
		end if
		strSQL = strSQL & ";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		if not(isNull(counter))then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,counter)
		end if
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objProductField
			do while not objRS.EOF
				Set objProductField = new ProductFieldClass
				strID = objRS("id_field")
				iCount = objRS("counter")
				objProductField.setID(strID)
				objProductField.setFoCounter(iCount)	
				objProductField.setIdOrder(objRS("id_order"))			
				objProductField.setIdProd(objRS("id_prod"))	
				objProductField.setQtaProd(objRS("qta_prod"))
				objProductField.setTypeField(objRS("type"))	
				objProductField.setSelValue(objRS("value"))	
				objProductField.setDescription(objRS("desc"))	
				
				if(objDict.Exists(iCount)) then
					Set objInnerDict = objDict(iCount)
					objInnerDict.add objProductField,""
					
					objDict.add iCount, objInnerDict
					Set objInnerDict = nothing
				else
					Set objInnerDict = Server.CreateObject("Scripting.Dictionary")
					objInnerDict.add objProductField,""
					
					objDict.add iCount, objInnerDict	
					Set objInnerDict = nothing
				end if
					
				Set objProductField = nothing
				objRS.moveNext()
			loop
							
			Set findListFieldXOrderByProd = objDict			
			Set objDict = nothing	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertFieldXOrder(counter, idOrder, idProd, idField, qtaProd, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO product_fields_x_order(`counter`, id_order, id_prod, id_field, qta_prod, `value`) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qtaProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,value)
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
		
	Public Sub modifyFieldXOrder(counter, idOrder, idField, idProd, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE product_fields_x_order SET "
		strSQL = strSQL & "qta_prod=?,"
		strSQL = strSQL & "`value`=?"
		strSQL = strSQL & " WHERE `counter`=?, id_order=? AND id_prod=? AND id_field=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qtaProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,value)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
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
		
	Public Sub deleteFieldXOrder(idOrder, idProd, idField, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_x_order WHERE id_order=? AND id_prod=? AND id_field=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
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
		
	Public Sub deleteFieldXOrderNoTransaction(idOrder, idProd, idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_x_order WHERE id_order=? AND id_prod=? AND id_field=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteFieldXOrderByProd(counterProd, idOrder, idProd, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_x_order WHERE `counter`=? AND id_order=? AND id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counterProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
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
		
	Public Sub deleteFieldXOrderByProdNoTransaction(counterProd, idOrder, idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_x_order WHERE `counter`=? AND id_order=? AND id_prod=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counterProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		

	'************************* GESTIONE FIELD X CARRELLO *******************************
		
	Public Function findFieldXCard(counter, idCard, idProd, idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldXCard = null		
		strSQL = "SELECT * FROM product_fields_x_card WHERE `counter`=? AND id_card =? AND id_prod=? AND id_field=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,counter)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idCard)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objProductField = new ProductFieldClass
			strID = objRS("id_field")
			objProductField.setID(strID)
			objProductField.setFcCounter(objRS("counter"))	
			objProductField.setIdCard(objRS("id_card"))			
			objProductField.setIdProd(objRS("id_prod"))	
			objProductField.setQtaProd(objRS("qta_prod"))
			objProductField.setSelValue(objRS("value"))
			Set findFieldXCard = objProductField			
			Set objProductField = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findListFieldXCardByProd(counter, idCard, idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findListFieldXCardByProd = null		
		strSQL = "SELECT product_fields_x_card.*, product_fields.`type`, product_fields.description as `desc` FROM product_fields_x_card LEFT JOIN product_fields ON product_fields_x_card.id_field=product_fields.id"
		strSQL = strSQL & " WHERE id_card =? AND id_prod=?"
		if not(isNull(counter))then
			strSQL = strSQL & " AND counter=?"
		end if
		strSQL = strSQL & ";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idCard)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idProd)
		if not(isNull(counter))then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,counter)
		end if
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objProductField
			do while not objRS.EOF
				Set objProductField = new ProductFieldClass
				strID = objRS("id_field")
				iCount = objRS("counter")
				objProductField.setID(strID)
				objProductField.setFcCounter(iCount)	
				objProductField.setIdCard(objRS("id_card"))			
				objProductField.setIdProd(objRS("id_prod"))	
				objProductField.setQtaProd(objRS("qta_prod"))
				objProductField.setTypeField(objRS("type"))	
				objProductField.setSelValue(objRS("value"))	
				objProductField.setDescription(objRS("desc"))	
				
				if(objDict.Exists(iCount)) then
					Set objInnerDict = objDict(iCount)
					objInnerDict.add objProductField,""
					
					objDict.add iCount, objInnerDict
					Set objInnerDict = nothing
				else
					Set objInnerDict = Server.CreateObject("Scripting.Dictionary")
					objInnerDict.add objProductField,""
					
					objDict.add iCount, objInnerDict	
					Set objInnerDict = nothing
				end if
					
				Set objProductField = nothing
				objRS.moveNext()
			loop
							
			Set findListFieldXCardByProd = objDict			
			Set objDict = nothing	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertFieldXCard(counter, idCard, idProd, idField, qtaProd, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO product_fields_x_card(`counter`, id_card, id_prod, id_field, qta_prod, `value`) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idCard)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qtaProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,value)
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
		
	Public Sub modifyFieldXCard(counter, idCard, idField, idProd, qtaProd, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE product_fields_x_card SET "
		strSQL = strSQL & "qta_prod=?,"
		strSQL = strSQL & "`value`=?"
		strSQL = strSQL & " WHERE `counter`=? AND id_card=? AND id_prod=? AND id_field=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qtaProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,value)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idCard)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
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
		
	Public Sub deleteFieldXCard(idCard, idProd, idField, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_x_card WHERE id_card=? AND id_prod=? AND id_field=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idCard)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
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
		
	Public Sub deleteFieldXCardNoTransaction(idCard, idProd, idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_x_card WHERE id_card=? AND id_prod=? AND id_field=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idCard)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteFieldXCardByProd(counterProd, idCard, idProd, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM product_fields_x_card WHERE `counter`=? AND id_card=? AND id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counterProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idCard)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
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
		
	Public Sub deleteFieldXCardByProdNoTransaction(counterProd, idCard, idProd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM product_fields_x_card WHERE `counter`=? AND id_card=? AND id_prod=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counterProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idCard)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idProd)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
	

	'************************* GESTIONE FIELD RENDERING *******************************
	
	Public Function renderProductFieldHTML(productField,cssClass, customFieldPrefix, idProd, defaultMatchValue, translator, isClient, isEditable)	
		Dim fieldMatchValue, spitValues, keyPress, maxLenght, style
		
		fieldMatchValue = defaultMatchValue
		
		'if not(idProd="") then
		'	on error resume next
		'		Set fieldMatchValue = findFieldMatch(productField.getID(),idProd)
		'		if (Instr(1, typename(fieldMatchValue), "dictionary", 1) > 0) then
		'			fieldMatchValue = fieldMatchValue.Item("value")
		'		end if
		'	if Err.number <> 0 then
				''response.write(Err.description)
		'	end if			
		'end if
		
		on error resume next			
		keyPress = ""
		select Case productField.getTypeContent()
		Case 2		
			keyPress = " onkeypress=""javascript:return isInteger(event);"""
		Case 3		
			keyPress = " onkeypress=""javascript:return isDouble(event);"""
		Case Else
		End Select		
		
		maxLenght = ""		
		if not(productField.getMaxLenght()="") AND (productField.getMaxLenght()>0) then
			maxLenght = " maxlength="""&productField.getMaxLenght()&""""
		end if
		
		style = ""		
		if not(cssClass="") then
			style = " class="""&cssClass&""""
		end if

		if isNull(customFieldPrefix) then
			customFieldPrefix = ""
		end if
		
		renderProductFieldHTML = ""		

		select Case productField.getTypeField()
		Case 1
			if(Trim(fieldMatchValue)<>"")then fieldMatchValue=Server.HTMLEncode(fieldMatchValue) end if
			
			if(isClient) then
				if(isEditable)then 
					'renderProductFieldHTML = "<input type=""text"" name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" value="""&fieldMatchValue&""" "&style&" "&keyPress&maxLenght&" onfocus=""cleanInputField('"&getFieldPrefix()&customFieldPrefix&productField.getID()&"');"" onBlur=""restoreInputField('"&getFieldPrefix()&customFieldPrefix&productField.getID()&"','"&fieldMatchValue&"');""/>"
					renderProductFieldHTML = "<input type=""text"" name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" value="""&fieldMatchValue&""" "&style&" "&keyPress&maxLenght&"/>"
	
					if(CInt(productField.getTypeContent())=4) then
						renderProductFieldHTML = renderProductFieldHTML & "<script>"
						renderProductFieldHTML = renderProductFieldHTML & "$(function() {"
							renderProductFieldHTML = renderProductFieldHTML & "$('#"&getFieldPrefix()&customFieldPrefix&productField.getID()&"').datepicker({"
								renderProductFieldHTML = renderProductFieldHTML & "dateFormat: 'dd/mm/yy',"
								renderProductFieldHTML = renderProductFieldHTML & "changeMonth: true,"
								renderProductFieldHTML = renderProductFieldHTML & "changeYear: true"
								'renderProductFieldHTML = renderProductFieldHTML & ",yearRange: '1900:"&DatePart("yyyy",Date())&"'" 
							renderProductFieldHTML = renderProductFieldHTML & "});"
						renderProductFieldHTML = renderProductFieldHTML & "});"
						renderProductFieldHTML = renderProductFieldHTML & "</script>"
					end if					
				else
					renderProductFieldHTML = fieldMatchValue					
				end if
			else
				if(isEditable)then
					renderProductFieldHTML = fieldMatchValue
				else		
					'renderProductFieldHTML = "<input type=""text"" name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" value="""&fieldMatchValue&""" "&style&" "&keyPress&maxLenght&" onfocus=""cleanInputField('"&getFieldPrefix()&customFieldPrefix&productField.getID()&"');"" onBlur=""restoreInputField('"&getFieldPrefix()&customFieldPrefix&productField.getID()&"','"&fieldMatchValue&"');""/>"		
					renderProductFieldHTML = "<input type=""text"" name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" value="""&fieldMatchValue&""" "&style&" "&keyPress&maxLenght&"/>"
	
					if(CInt(productField.getTypeContent())=4) then
						renderProductFieldHTML = renderProductFieldHTML & "<script>"
						renderProductFieldHTML = renderProductFieldHTML & "$(function() {"
							renderProductFieldHTML = renderProductFieldHTML & "$('#"&getFieldPrefix()&customFieldPrefix&productField.getID()&"').datepicker({"
								renderProductFieldHTML = renderProductFieldHTML & "dateFormat: 'dd/mm/yy',"
								renderProductFieldHTML = renderProductFieldHTML & "changeMonth: true,"
								renderProductFieldHTML = renderProductFieldHTML & "changeYear: true"
								'renderProductFieldHTML = renderProductFieldHTML & ",yearRange: '1900:"&DatePart("yyyy",Date())&"'" 
							renderProductFieldHTML = renderProductFieldHTML & "});"
						renderProductFieldHTML = renderProductFieldHTML & "});"
						renderProductFieldHTML = renderProductFieldHTML & "</script>"
					end if			
				end if
			end if
		Case 2
			if(Trim(fieldMatchValue)<>"")then fieldMatchValue=Server.HTMLEncode(fieldMatchValue) end if
			
			if(isClient) then
				if(isEditable)then
					renderProductFieldHTML = "<textarea name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
				else
					renderProductFieldHTML = fieldMatchValue
				end if
			else
				if(isEditable)then
					renderProductFieldHTML = fieldMatchValue
				else
					renderProductFieldHTML = "<textarea name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
				end if
			end if
		Case 3
			renderProductFieldHTML = "<select name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" "&style&">"
			renderProductFieldHTML = renderProductFieldHTML & "<option value=""""></option>"			
			
			
			On Error Resume Next
			'spitValues = Split(productField.getValues(),",")
			spitValues = getListProductFieldValues(productField.getID()).Keys
			for each x in spitValues
				'************* TODO: VERIFICARE QUESTA PARTE DI CHIAMATE AL DB, SERVE PER ESCLUDERE GLI ELEMENTI DELLA LISTA CHE HANNO QTA < 1
				'************* OGNI CHIAMATA CON LA TERNA idField, idProd, strValue FA UNA RICHIESTA AL DB;
				'************* SE DIVENTA LENTA LA VISUALIZZAZIONE DELLE COMBO COMMENTARE LE SEI RIGHE SEGUENTI	E L'END IF
				bolXQta = true	
				xQta = findFieldValueMatch(productField.getID(), idProd, Trim(x))
				if(Trim(xQta) <> "" AND not(isNull(xQta)))then
					if(Cint(xQta)<=0)then
						bolXQta = false
					end if
				end if
				if(bolXQta)then			
					selected = ""
					if (strComp(Trim(x), fieldMatchValue, 1) = 0) then selected=" selected" end if
					label= Trim(x)
					if not(translator.getTranslated("portal.commons.product_field.label."&label)="") then label=translator.getTranslated("portal.commons.product_field.label."&label) end if
					renderProductFieldHTML = renderProductFieldHTML & "<OPTION VALUE="""&x&""""&selected&">"&label&"</OPTION>"
				end if
			next
			if(Err.number <> 0)then
			end if
			
			renderProductFieldHTML = renderProductFieldHTML & "</select>"
		Case 4
			renderProductFieldHTML ="<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&customFieldPrefix&productField.getID()&""">"
			
			renderProductFieldHTML = renderProductFieldHTML & "<select name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" multiple size="""&productField.getMaxLenght()&""" "&style&">"
			renderProductFieldHTML = renderProductFieldHTML & "<option value=""""></option>"			
			
			On Error Resume Next
			'spitValues = Split(productField.getValues(),",")
			spitValues = getListProductFieldValues(productField.getID()).Keys
			for each x in spitValues
				'************* TODO: VERIFICARE QUESTA PARTE DI CHIAMATE AL DB, SERVE PER ESCLUDERE GLI ELEMENTI DELLA LISTA CHE HANNO QTA < 1
				'************* OGNI CHIAMATA CON LA TERNA idField, idProd, strValue FA UNA RICHIESTA AL DB;
				'************* SE DIVENTA LENTA LA VISUALIZZAZIONE DELLE COMBO COMMENTARE LE SEI RIGHE SEGUENTI	E L'END IF	
				bolXQta = true	
				xQta = findFieldValueMatch(productField.getID(), idProd, Trim(x))
				if(Trim(xQta) <> "" AND not(isNull(xQta)))then
					if(Cint(xQta)<=0)then
						bolXQta = false
					end if
				end if
				if(bolXQta)then
					selected = ""
					if (strComp(Trim(x), fieldMatchValue, 1) = 0) then selected=" selected" end if
					label= Trim(x)
					if not(translator.getTranslated("portal.commons.product_field.label."&label)="") then label=translator.getTranslated("portal.commons.product_field.label."&label) end if
					renderProductFieldHTML = renderProductFieldHTML & "<OPTION VALUE="""&x&""""&selected&">"&label&"</OPTION>"
				end if
			next
			
			renderProductFieldHTML = renderProductFieldHTML & "</select>"
		Case 5			
			on error Resume Next
			'if not(productField.getValues() = "") then
				'spitValues = Split(productField.getValues(),",")
			'end if
			spitValues = getListProductFieldValues(productField.getID()).Keys
			K=1		
	
			renderProductFieldHTML =renderProductFieldHTML & "<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&customFieldPrefix&productField.getID()&""">"

			for each y in spitValues
				checked = ""
				if not(fieldMatchValue = "") then
					spitMatchValues = Split(fieldMatchValue,",")
					for j=0 to Ubound(spitMatchValues)
						if(strComp(Trim(spitMatchValues(j)), Trim(y), 1) = 0) then
							checked=" checked='checked'"
							exit for
						end if
					next
				end if
				newLine = ""
				if((k Mod 4) = 0) then newLine="<br/>" end if
				label= Trim(y)
				if not(translator.getTranslated("portal.commons.product_field.label."&label)="") then label=translator.getTranslated("portal.commons.product_field.label."&label) end if
				renderProductFieldHTML =renderProductFieldHTML & label&"&nbsp;<input type=""checkbox"" "&style&" value="""&y&""" name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" "&checked&"/>"&newLine
				k = k+1
			next
			
			if Err.number <> 0 then
				'response.write(Err.description)
			end if
		Case 6			
			on error Resume Next
			'if not(productField.getValues() = "") then
				'spitValues = Split(productField.getValues(),",")
			'end if
			spitValues = getListProductFieldValues(productField.getID()).Keys
			K=1
			
			renderProductFieldHTML =renderProductFieldHTML & "<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&customFieldPrefix&productField.getID()&""">"

			for each y in spitValues
				checked = ""
				if not(fieldMatchValue = "") then
					if (strComp(fieldMatchValue, Trim(y), 1) = 0) then checked=" checked='checked'" end if
				end if
				newLine = ""
				if((k Mod 4) = 0) then newLine="<br/>" end if
				label= Trim(y)
				if not(translator.getTranslated("portal.commons.product_field.label."&label)="") then label=translator.getTranslated("portal.commons.product_field.label."&label) end if
				renderProductFieldHTML =renderProductFieldHTML & label&"&nbsp;<input type=""radio"" "&style&" value="""&y&""" name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" "&checked&"/>"&newLine
				k = k+1
			next
			
			if Err.number <> 0 then
				'response.write(Err.description)
			end if
		Case 7
			if(Trim(fieldMatchValue)<>"")then fieldMatchValue=Server.HTMLEncode(fieldMatchValue) end if
			
			renderProductFieldHTML = "<input type=""hidden"" value="""&fieldMatchValue&""" name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&"""/>"
		Case 8
			renderProductFieldHTML = "<input type=""file"" name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" "&style&" />"
		Case 9
			if(Trim(fieldMatchValue)<>"")then fieldMatchValue=Server.HTMLEncode(fieldMatchValue) end if
			
			if(isClient) then
				if(isEditable)then
					renderProductFieldHTML = "<script type=""text/javascript"">"
					'renderProductFieldHTML = renderProductFieldHTML & "$(document).ready(function() {$(""#"&getFieldPrefix()&customFieldPrefix&productField.getID()&""").cleditor(cloptions"&getFieldPrefix()&customFieldPrefix&productField.getID()&");});"
					renderProductFieldHTML = renderProductFieldHTML & "$.cleditor.defaultOptions.width = 280;"
					renderProductFieldHTML = renderProductFieldHTML & "$.cleditor.defaultOptions.height = 200;"
					renderProductFieldHTML = renderProductFieldHTML & "$.cleditor.defaultOptions.controls = ""bold italic underline strikethrough subscript superscript | font size style | color highlight removeformat | bullets numbering | alignleft center alignright justify | rule | cut copy paste | image"";"		
					renderProductFieldHTML = renderProductFieldHTML & "$(document).ready(function() {$(""#"&getFieldPrefix()&customFieldPrefix&productField.getID()&""").cleditor();});"
					renderProductFieldHTML = renderProductFieldHTML & "</script>"
					renderProductFieldHTML = renderProductFieldHTML & "<textarea name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
				else
					renderProductFieldHTML = fieldMatchValue
				end if
			else
				if(isEditable)then
					renderProductFieldHTML = fieldMatchValue
				else
					renderProductFieldHTML = "<script type=""text/javascript"">"
					'renderProductFieldHTML = renderProductFieldHTML & "$(document).ready(function() {$(""#"&getFieldPrefix()&customFieldPrefix&productField.getID()&""").cleditor(cloptions"&getFieldPrefix()&customFieldPrefix&productField.getID()&");});"
					renderProductFieldHTML = renderProductFieldHTML & "$.cleditor.defaultOptions.width = 280;"
					renderProductFieldHTML = renderProductFieldHTML & "$.cleditor.defaultOptions.height = 200;"
					renderProductFieldHTML = renderProductFieldHTML & "$.cleditor.defaultOptions.controls = ""bold italic underline strikethrough subscript superscript | font size style | color highlight removeformat | bullets numbering | alignleft center alignright justify | rule | cut copy paste | image"";"		
					renderProductFieldHTML = renderProductFieldHTML & "$(document).ready(function() {$(""#"&getFieldPrefix()&customFieldPrefix&productField.getID()&""").cleditor();});"
					renderProductFieldHTML = renderProductFieldHTML & "</script>"
					renderProductFieldHTML = renderProductFieldHTML & "<textarea name="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&productField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
				end if
			end if				
		Case Else
		End Select	
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function
	
	Public Function renderProductFieldJS(productField, formName, customFieldPrefix, translator,defaultMatchValue, frontOrBack)	
		on error resume next

		if isNull(customFieldPrefix) then
			customFieldPrefix = ""
		end if
		
		renderProductFieldJS = ""	
		

		select Case productField.getTypeField()
		Case 1,2
			renderProductFieldJS = "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&"){"
			renderProductFieldJS = renderProductFieldJS & "var "&getFieldPrefix()&customFieldPrefix&productField.getID()&"_default_match_values = """&defaultMatchValue&""";"
			renderProductFieldJS = renderProductFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value.toLowerCase() == "&getFieldPrefix()&customFieldPrefix&productField.getID()&"_default_match_values.toLowerCase()){"
				renderProductFieldJS = renderProductFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value = """";"
			renderProductFieldJS = renderProductFieldJS &"}"				
		
			if(CInt(productField.getRequired())=1)then
				'se backoffice, verifico se è stata selezionata la checkbox del field e solo in quel caso attivo il controllo required
				if(frontOrBack="1")then
					'caso frontend, do nothing
				elseif(frontOrBack="2")then
					'caso backend
					renderProductFieldJS = renderProductFieldJS & "if(document."&formName&".prod_field_active != null){"
						renderProductFieldJS = renderProductFieldJS & "if(document."&formName&".prod_field_active.length == null){"
							renderProductFieldJS = renderProductFieldJS & "if(document."&formName&".prod_field_active.checked){"
								renderProductFieldJS = renderProductFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value == """"){"
									renderProductFieldJS = renderProductFieldJS &"alert("""&translator.getTranslated("portal.commons.product_field.js.alert.insert_"&productField.getDescription())&""");"
									renderProductFieldJS = renderProductFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".focus();"
									renderProductFieldJS = renderProductFieldJS &"return false;"
								renderProductFieldJS = renderProductFieldJS &"}"					
							renderProductFieldJS = renderProductFieldJS & "}"
						renderProductFieldJS = renderProductFieldJS & "}else{"
							renderProductFieldJS = renderProductFieldJS & "for(k=0; k<document."&formName&".prod_field_active.length; k++){"
								renderProductFieldJS = renderProductFieldJS & "if(document."&formName&".prod_field_active[k].checked && document."&formName&".prod_field_active[k].value=="""&productField.getID()&"-"&productField.getTypeField()&"""){"						
									renderProductFieldJS = renderProductFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value == """"){"
										renderProductFieldJS = renderProductFieldJS &"alert("""&translator.getTranslated("portal.commons.product_field.js.alert.insert_"&productField.getDescription())&""");"
										renderProductFieldJS = renderProductFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".focus();"
										renderProductFieldJS = renderProductFieldJS &"return false;"
									renderProductFieldJS = renderProductFieldJS &"}"
								renderProductFieldJS = renderProductFieldJS & "}"
							renderProductFieldJS = renderProductFieldJS & "}"
						renderProductFieldJS = renderProductFieldJS & "}"
					renderProductFieldJS = renderProductFieldJS & "}"
				end if
			end if

			if(CInt(productField.getTypeContent())=2) then
				renderProductFieldJS = renderProductFieldJS &"if(isNaN(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value)){"
					renderProductFieldJS = renderProductFieldJS &"alert("""&translator.getTranslated("portal.commons.product_field.js.alert.isnan_value")&""");"
					renderProductFieldJS = renderProductFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".focus();"
					renderProductFieldJS = renderProductFieldJS &"return false;"	
				renderProductFieldJS = renderProductFieldJS &"}"			
			end if

			if(CInt(productField.getTypeContent())=3) then
				renderProductFieldJS = renderProductFieldJS &"if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value.length > 0 && (!checkDoubleFormatExt(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value) || document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value.indexOf(""."")!=-1)){"
					renderProductFieldJS = renderProductFieldJS &"alert("""&translator.getTranslated("portal.commons.product_field.js.alert.isnan_value")&""");"
					renderProductFieldJS = renderProductFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".focus();"
					renderProductFieldJS = renderProductFieldJS &"return false;"
				renderProductFieldJS = renderProductFieldJS &"}"		
			end if
			renderProductFieldJS = renderProductFieldJS &"}"

		Case 3,4
			if(CInt(productField.getRequired())=1)then
				renderProductFieldJS = "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&"){"
					renderProductFieldJS = renderProductFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".options[document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".selectedIndex].value == """"){"
						renderProductFieldJS = renderProductFieldJS &"alert("""&translator.getTranslated("portal.commons.product_field.js.alert.insert_"&productField.getDescription())&""");"
						renderProductFieldJS = renderProductFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".focus();"
						renderProductFieldJS = renderProductFieldJS &"return false;"
					renderProductFieldJS = renderProductFieldJS &"}"
				renderProductFieldJS = renderProductFieldJS &"}"			
			end if		
		Case 5,6
			renderProductFieldJS = "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&"){"
			renderProductFieldJS = renderProductFieldJS & "var "&getFieldPrefix()&customFieldPrefix&productField.getID()&"_values = """";"
			renderProductFieldJS = renderProductFieldJS &"if (document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&"){"
				renderProductFieldJS = renderProductFieldJS &"if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&" != null){"				
					renderProductFieldJS = renderProductFieldJS &"if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".length == null){"
						renderProductFieldJS = renderProductFieldJS &"if (document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".checked){"
							renderProductFieldJS = renderProductFieldJS &getFieldPrefix()&customFieldPrefix&productField.getID()&"_values = "&getFieldPrefix()&customFieldPrefix&productField.getID()&"_values + document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".value + "","";"
							renderProductFieldJS = renderProductFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".checked=false;"
						renderProductFieldJS = renderProductFieldJS &"}"						
					renderProductFieldJS = renderProductFieldJS &"}else{"
						renderProductFieldJS = renderProductFieldJS &"for (var i=0; i < document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&".length; i++){"
							renderProductFieldJS = renderProductFieldJS &"if (document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&"[i].checked){"
								renderProductFieldJS = renderProductFieldJS &getFieldPrefix()&customFieldPrefix&productField.getID()&"_values = "&getFieldPrefix()&customFieldPrefix&productField.getID()&"_values + document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&"[i].value + "","";"
								renderProductFieldJS = renderProductFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&productField.getID()&"[i].checked=false;"
							renderProductFieldJS = renderProductFieldJS &"}"
						renderProductFieldJS = renderProductFieldJS &"}"						
					renderProductFieldJS = renderProductFieldJS &"}"
						
					renderProductFieldJS = renderProductFieldJS &getFieldPrefix()&customFieldPrefix&productField.getID()&"_values = "&getFieldPrefix()&customFieldPrefix&productField.getID()&"_values.substring(0, "&getFieldPrefix()&customFieldPrefix&productField.getID()&"_values.lastIndexOf(','));"
				renderProductFieldJS = renderProductFieldJS &"}"
			renderProductFieldJS = renderProductFieldJS &"}"
			renderProductFieldJS = renderProductFieldJS &"document."&formName&".hidden_"&getFieldPrefix()&customFieldPrefix&productField.getID()&".value = "&getFieldPrefix()&customFieldPrefix&productField.getID()&"_values;"
				
			if(CInt(productField.getRequired())=1)then							
				renderProductFieldJS = renderProductFieldJS &"if ("&getFieldPrefix()&customFieldPrefix&productField.getID()&"_values ==""""){"
				renderProductFieldJS = renderProductFieldJS &"alert("""&translator.getTranslated("portal.commons.product_field.js.alert.insert_"&productField.getDescription())&""");"
				renderProductFieldJS = renderProductFieldJS &"return false;"
				renderProductFieldJS = renderProductFieldJS &"}"	
			end if
			renderProductFieldJS = renderProductFieldJS &"}"
		Case Else
		End Select
 
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function
	
	Public Function getFieldPrefix()
		getFieldPrefix = "productfield"
	End Function
	
End Class
%>