<%
Class VoucherClass
	Private id
	Private voucher_type
	Private label
	Private description
	Private valore
	Private activate
	Private operation
	Private max_generation
	Private max_usage
	Private enable_date
	Private expire_date
	Private objVoucherCode
	Private excludeProdRule
	
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub	
	
	Public Function getVoucherType()
		getVoucherType = voucher_type
	End Function
	
	Public Sub setVoucherType(strVoucherType)
		voucher_type = strVoucherType
	End Sub
	
	Public Function getLabel()
		getLabel = label
	End Function
	
	Public Sub setLabel(strLabel)
		label = strLabel
	End Sub
	
	Public Function getDescrizione()
		getDescrizione = description
	End Function
	
	Public Sub setDescrizione(strDesc)
		description = strDesc
	End Sub		

	Public Function getValore()
		getValore = valore
	End Function
	
	Public Sub setValore(strValore)
		valore = strValore
	End Sub
	
	Public Function getActivate()
		getActivate = activate
	End Function
	
	Public Sub setActivate(strActivate)
		activate = strActivate
	End Sub	

	Public Function getOperation()
		getOperation = operation
	End Function
	
	Public Sub setOperation(strOperation)
		operation = strOperation
	End Sub

	Public Function getMaxGeneration()
		getMaxGeneration = max_generation
	End Function
	
	Public Sub setMaxGeneration(strMaxGeneration)
		max_generation = strMaxGeneration
	End Sub

	Public Function getMaxUsage()
		getMaxUsage = max_usage
	End Function
	
	Public Sub setMaxUsage(strMaxUsage)
		max_usage = strMaxUsage
	End Sub

	Public Function getEnableDate()
		getEnableDate = enable_date
	End Function
	
	Public Sub setEnableDate(strEnableDate)
		enable_date = strEnableDate
	End Sub

	Public Function getExpireDate()
		getExpireDate = expire_date
	End Function
	
	Public Sub setExpireDate(strExpireDate)
		expire_date = strExpireDate
	End Sub

	Public Function getObjVoucherCode()
		Set getObjVoucherCode = objVoucherCode
	End Function
	
	Public Sub setObjVoucherCode(strObjVoucherCode)
		Set objVoucherCode = strObjVoucherCode
	End Sub

	Public Function getExcludeProdRule()
		getExcludeProdRule = excludeProdRule
	End Function
	
	Public Sub setExcludeProdRule(strExcludeProdRule)
		excludeProdRule = strExcludeProdRule
	End Sub
	

	Public Function validateVoucherCode(voucher_code)
		validateVoucherCode = null	
		
		on error resume next
		Set objVoucherExt = findExtendedVoucherByCode(voucher_code)		
		'response.write("typename(objVoucherExt): "&typename(objVoucherExt)&"<br>")
		
		if not(isNull(objVoucherExt)) then
			if(objVoucherExt.getActivate())then
				vtype=objVoucherExt.getVoucherType()				
				'response.write("vtype: "&vtype&"<br>")		
				Select Case CInt(vtype)
				Case 0
					'response.write("objVoucherExt.getObjVoucherCode().getUsageCounter(): "&objVoucherExt.getObjVoucherCode().getUsageCounter()&"<br>")
					'response.write("if: "& (CLng(objVoucherExt.getObjVoucherCode().getUsageCounter())=0)&"<br>")
					if(CLng(objVoucherExt.getObjVoucherCode().getUsageCounter())=0)then
						Set validateVoucherCode = objVoucherExt
					end if
				Case 1
					if(CLng(objVoucherExt.getObjVoucherCode().getUsageCounter())<Clng(objVoucherExt.getMaxUsage()) OR Clng(objVoucherExt.getMaxUsage())=-1)then
						Set validateVoucherCode = objVoucherExt
					end if					
				Case 2
					tmpEnableDate = objVoucherExt.getEnableDate()
					tmpExpireDate = objVoucherExt.getExpireDate()
					tmpInsertDate = objVoucherExt.getObjVoucherCode().getInsertDate()
					'response.write("if: "& (CLng(objVoucherExt.getObjVoucherCode().getUsageCounter())=0)&"<br>")
					'response.write("tmpEnableDate: "&tmpEnableDate&"<br>")
					'response.write("tmpExpireDate: "&tmpExpireDate&"<br>")
					'response.write("tmpInsertDate: "&tmpInsertDate&"<br>")
					if(CLng(objVoucherExt.getObjVoucherCode().getUsageCounter())=0 AND DateDiff("d",tmpEnableDate,tmpInsertDate)>=0 AND DateDiff("d",tmpInsertDate,tmpExpireDate)>=0)then
						Set validateVoucherCode = objVoucherExt
					end if
				Case 3
					tmpEnableDate = objVoucherExt.getEnableDate()
					tmpExpireDate = objVoucherExt.getExpireDate()
					tmpInsertDate = objVoucherExt.getObjVoucherCode().getInsertDate()
					if((CLng(objVoucherExt.getObjVoucherCode().getUsageCounter())<Clng(objVoucherExt.getMaxUsage()) OR Clng(objVoucherExt.getMaxUsage())=-1) AND DateDiff("d",tmpEnableDate,tmpInsertDate)>=0 AND DateDiff("d",tmpInsertDate,tmpExpireDate)>=0)then
						Set validateVoucherCode = objVoucherExt
					end if
					
				Case 4
					if(CLng(objVoucherExt.getObjVoucherCode().getUsageCounter())=0 AND not(isNull(objVoucherExt.getObjVoucherCode().getIdUserRef())) AND objVoucherExt.getObjVoucherCode().getIdUserRef()<>"")then
						Set validateVoucherCode = objVoucherExt
					end if					
				Case else
				End Select
			end if
		end if
		
		if Err.number <> 0 then
			validateVoucherCode = null
		end if		
	End Function


	Public Function getCampaignList(voucher_type, activate)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getCampaignList = null		
		strSQL = "SELECT * FROM voucher_campaign"
		
		if (isNull(voucher_type) AND isNull(activate)) then
			strSQL = "SELECT * FROM voucher_campaign"
		else
			strSQL = strSQL & " WHERE"
			if not(isNull(voucher_type)) then strSQL = strSQL & " AND voucher_type IN("&voucher_type&")"
			if not(isNull(activate)) then strSQL = strSQL & " AND activate=?"
		end if
		
		strSQL = strSQL & " ORDER BY voucher_type ASC, label ASC;"
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if (isNull(activate)) then
		else
			if not(isNull(activate)) then objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activate)			
		end if
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objVoucher
			do while not objRS.EOF				
				Set objVoucher = new VoucherClass
				strID = objRS("id")
				objVoucher.setID(strID)
				objVoucher.setVoucherType(objRS("voucher_type"))
				objVoucher.setLabel(objRS("label"))	
				objVoucher.setDescrizione(objRS("description"))	
				objVoucher.setActivate(objRS("activate"))		
				objVoucher.setValore(objRS("valore"))	
				objVoucher.setOperation(objRS("operation"))	
				objVoucher.setMaxGeneration(objRS("max_generation"))	
				objVoucher.setMaxUsage(objRS("max_usage"))	
				objVoucher.setEnableDate(objRS("enable_date"))	
				objVoucher.setExpireDate(objRS("expire_date"))	
				objVoucher.setExcludeProdRule(objRS("exclude_prod_rule"))
				objDict.add strID, objVoucher
				Set objVoucher = nothing
				objRS.moveNext()
			loop							
			Set getCampaignList = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	
	Public Function findCampaignByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		findCampaignByID = null		
		strSQL = "SELECT * FROM voucher_campaign WHERE id =?;"
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
			Dim objVoucher		
			Set objVoucher = new VoucherClass
			strID = objRS("id")
			objVoucher.setID(strID)
			objVoucher.setVoucherType(objRS("voucher_type"))
			objVoucher.setLabel(objRS("label"))	
			objVoucher.setDescrizione(objRS("description"))	
			objVoucher.setActivate(objRS("activate"))		
			objVoucher.setValore(objRS("valore"))	
			objVoucher.setOperation(objRS("operation"))	
			objVoucher.setMaxGeneration(objRS("max_generation"))	
			objVoucher.setMaxUsage(objRS("max_usage"))	
			objVoucher.setEnableDate(objRS("enable_date"))	
			objVoucher.setExpireDate(objRS("expire_date"))
			objVoucher.setExcludeProdRule(objRS("exclude_prod_rule"))					
			Set findCampaignByID = objVoucher			
			Set objRule = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findCampaignByLabel(label)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		findCampaignByLabel = null		
		strSQL = "SELECT * FROM voucher_campaign WHERE LOWER(label)=?;"
		strSQL = Trim(strSQL)
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,LCase(label))
		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then			
			Dim objVoucher		
			Set objVoucher = new VoucherClass
			strID = objRS("id")
			objVoucher.setID(strID)
			objVoucher.setVoucherType(objRS("voucher_type"))
			objVoucher.setLabel(objRS("label"))	
			objVoucher.setDescrizione(objRS("description"))	
			objVoucher.setActivate(objRS("activate"))		
			objVoucher.setValore(objRS("valore"))	
			objVoucher.setOperation(objRS("operation"))	
			objVoucher.setMaxGeneration(objRS("max_generation"))	
			objVoucher.setMaxUsage(objRS("max_usage"))	
			objVoucher.setEnableDate(objRS("enable_date"))	
			objVoucher.setExpireDate(objRS("expire_date"))
			objVoucher.setExcludeProdRule(objRS("exclude_prod_rule"))					
			Set findCampaignByLabel = objVoucher			
			Set objRule = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function insertCampaign(label, voucher_type, description, valore, operation, activate, max_generation, max_usage, enable_date, expire_date, exclude_prod_rule, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		insertCampaign = -1
		
		strSQL = "INSERT INTO voucher_campaign(label, voucher_type, description, valore, operation, activate, max_generation, max_usage, enable_date, expire_date, exclude_prod_rule) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		if(isNull(enable_date) OR enable_date = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		if(isNull(expire_date) OR expire_date = "") then
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
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,voucher_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,description)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,operation)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activate)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,max_generation)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,max_usage)
		if not isNull(enable_date) AND not(enable_date = "") then			
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,enable_date)
		end if
		if not isNull(expire_date) AND not(expire_date = "") then			
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,expire_date)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,exclude_prod_rule)
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(business_rules.id) as id FROM business_rules")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertCampaign = objRS("id")	
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
		
	Public Sub modifyCampaign(id, label, voucher_type, description, valore, operation, activate, max_generation, max_usage, enable_date, expire_date, exclude_prod_rule, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE voucher_campaign SET "
		strSQL = strSQL & "label=?,"
		strSQL = strSQL & "voucher_type=?,"
		strSQL = strSQL & "description=?,"
		strSQL = strSQL & "valore=?,"
		strSQL = strSQL & "operation=?,"
		strSQL = strSQL & "activate=?,"
		strSQL = strSQL & "max_generation=?,"
		strSQL = strSQL & "max_usage=?,"
		if(isNull(enable_date) OR enable_date = "") then
			strSQL = strSQL & "enable_date=NULL,"
		else
			strSQL = strSQL & "enable_date=?,"
		end if
		if(isNull(expire_date) OR expire_date = "") then
			strSQL = strSQL & "expire_date=NULL,"
		else
			strSQL = strSQL & "expire_date=?,"
		end if
		strSQL = strSQL & "exclude_prod_rule=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,voucher_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,description)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,operation)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activate)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,max_generation)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,max_usage)
		if not isNull(enable_date) AND not(enable_date = "") then			
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,enable_date)
		end if
		if not isNull(expire_date) AND not(expire_date = "") then			
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,expire_date)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,exclude_prod_rule)
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

	Public Sub deleteCampaign(id)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL2 = "DELETE FROM voucher_code WHERE id_voucher=?;"
		strSQL = "DELETE FROM voucher_campaign WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQL2
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,3,1,,id)
		
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

'***************************************************** VOUCHER ORDER ASSOCIATION METHODS

	Public Function findVoucherOrderAssociationsByCode(voucher_code)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		findVoucherOrderAssociationsByCode = null	
		strSQL = "SELECT * FROM voucher_x_ordine WHERE voucher_code=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,voucher_code)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then							
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objCode
			do while not objRS.EOF				
				Set objCode = new VoucherCodeClass
				strID = objRS("id_voucher")
				strOrderID = objRS("id_order")
				strCode = objRS("voucher_code")
				objCode.setID(strID)
				objCode.setOrderID(strOrderID)	
				objCode.setVoucherCode(strCode)	
				objCode.setValore(objRS("valore"))	
				objCode.setInsertDate(objRS("insert_date"))
				objDict.add strID&"-"&strOrderID&"-"&strCode, objCode
				Set objCode = nothing
				objRS.moveNext()
			loop							
			Set findVoucherOrderAssociationsByCode = objDict			
			Set objDict = nothing			
		end if	
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function findVoucherOrderAssociationsByOrder(id_order)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		findVoucherOrderAssociationsByOrder = null	
		strSQL = "SELECT * FROM voucher_x_ordine WHERE id_order=? ORDER BY insert_date;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then							
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objCode
			do while not objRS.EOF				
				Set objCode = new VoucherCodeClass
				strID = objRS("id_voucher")
				strOrderID = objRS("id_order")
				strCode = objRS("voucher_code")
				objCode.setID(strID)
				objCode.setOrderID(strOrderID)	
				objCode.setVoucherCode(strCode)	
				objCode.setValore(objRS("valore"))	
				objCode.setInsertDate(objRS("insert_date"))
				objDict.add strID&"-"&strOrderID&"-"&strCode, objCode
				Set objCode = nothing
				objRS.moveNext()
			loop							
			Set findVoucherOrderAssociationsByOrder = objDict			
			Set objDict = nothing			
		end if	
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function findLastVoucherOrderAssociationsByOrder(id_order)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		findLastVoucherOrderAssociationsByOrder = null	
		strSQL = "SELECT * FROM voucher_x_ordine WHERE id_order=? ORDER BY insert_date;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then							
			'Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objCode
			'do while not objRS.EOF				
				Set objCode = new VoucherCodeClass
				strID = objRS("id_voucher")
				strOrderID = objRS("id_order")
				strCode = objRS("voucher_code")
				objCode.setID(strID)
				objCode.setOrderID(strOrderID)	
				objCode.setVoucherCode(strCode)	
				objCode.setValore(objRS("valore"))	
				objCode.setInsertDate(objRS("insert_date"))
				'objDict.add strID&"-"&strOrderID&"-"&strCode, objCode
				'Set objCode = nothing
				'objRS.moveNext()
			'loop							
			'Set findVoucherOrderAssociationsByOrder = objDict			
			'Set objDict = nothing
			Set findLastVoucherOrderAssociationsByOrder = objCode
			Set objCode = nothing			
		end if	
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Sub insertVoucherOrder(id_order, voucher_code, id_voucher, valore, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO voucher_x_ordine(id_order, voucher_code, id_voucher, valore, insert_date) VALUES("
		strSQL = strSQL & "?,?,?,?,?);"
							
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,voucher_code)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_voucher)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,now())
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
		
	Public Sub modifyVoucherOrder(id_order, voucher_code, id_voucher, valore, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE voucher_x_ordine SET "
		strSQL = strSQL & "id_order=?,"
		strSQL = strSQL & "voucher_code=?,"
		strSQL = strSQL & "id_voucher=?,"
		strSQL = strSQL & "valore=?,"
		strSQL = strSQL & "insert_date=?"
		strSQL = strSQL & " WHERE id_order=? AND voucher_code=? AND id_voucher=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,voucher_code)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_voucher)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,now())
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,voucher_code)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_voucher)
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

	Public Sub deleteVoucherOrderByOrderID(id_order, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM voucher_x_ordine WHERE id_order=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
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

'***************************************************** VOUCHER CODE METHODS

	Public Function getListaVoucherCode(id_campaign)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getListaVoucherCode = null		
		strSQL = "SELECT * FROM voucher_code WHERE voucher_campaign=? ORDER BY id DESC;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_campaign)
		Set objRS = objCommand.Execute()

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objVCode
			do while not objRS.EOF				
				Set objVCode = new VoucherCodeClass
				strID = objRS("id")
				objVCode.setID(strID)
				objVCode.setVoucherCode(objRS("code"))
				objVCode.setVoucherCampaign(objRS("voucher_campaign"))
				objVCode.setInsertDate(objRS("insert_date"))	
				objVCode.setUsageCounter(objRS("usage_counter"))
				objVCode.setIdUserRef(objRS("id_user_ref"))	
				objDict.add strID, objVCode
				objRS.moveNext()
			loop
			Set objVCode = nothing							
			Set getListaVoucherCode = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findVoucherCodeByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		findVoucherCodeByID = null		
		strSQL = "SELECT * FROM voucher_code WHERE id =?;"
		
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
			Dim objVCode		
			Set objVCode = new VoucherCodeClass
			strID = objRS("id")
			objVCode.setID(strID)
			objVCode.setVoucherCode(objRS("code"))
			objVCode.setVoucherCampaign(objRS("voucher_campaign"))
			objVCode.setInsertDate(objRS("insert_date"))	
			objVCode.setUsageCounter(objRS("usage_counter"))
			objVCode.setIdUserRef(objRS("id_user_ref"))								
			Set findVoucherCodeByID = objVCode			
			Set objVCode = nothing					
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findVoucherCodeByCode(voucher_code)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		findVoucherCodeByCode = null		
		strSQL = "SELECT * FROM voucher_code WHERE code=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,voucher_code)
		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then			
			Dim objVCode		
			Set objVCode = new VoucherCodeClass
			strID = objRS("id")
			objVCode.setID(strID)
			objVCode.setVoucherCode(objRS("code"))
			objVCode.setVoucherCampaign(objRS("voucher_campaign"))
			objVCode.setInsertDate(objRS("insert_date"))	
			objVCode.setUsageCounter(objRS("usage_counter"))
			objVCode.setIdUserRef(objRS("id_user_ref"))								
			Set findVoucherCodeByCode = objVCode			
			Set objVCode = nothing					
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findExtendedVoucherByCode(voucher_code)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		findExtendedVoucherByCode = null		
		strSQL = "SELECT voucher_code.id as vcid, code, voucher_campaign, insert_date, usage_counter, id_user_ref, "
		strSQL = strSQL&" voucher_campaign.*"
		strSQL = strSQL&" FROM voucher_code LEFT JOIN voucher_campaign ON voucher_code.voucher_campaign=voucher_campaign.id WHERE code=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,voucher_code)
		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then			
			Dim objVoucher		
			Set objVoucher = new VoucherClass
			Set objVC = new VoucherCodeClass
			
			objVC.setID(objRS("vcid"))
			objVC.setVoucherCode(objRS("code"))
			objVC.setVoucherCampaign(objRS("voucher_campaign"))
			objVC.setInsertDate(objRS("insert_date"))	
			objVC.setUsageCounter(objRS("usage_counter"))
			objVC.setIdUserRef(objRS("id_user_ref"))
			
			objVoucher.setID(objRS("id"))
			objVoucher.setVoucherType(objRS("voucher_type"))
			objVoucher.setLabel(objRS("label"))	
			objVoucher.setDescrizione(objRS("description"))	
			objVoucher.setActivate(objRS("activate"))		
			objVoucher.setValore(objRS("valore"))	
			objVoucher.setOperation(objRS("operation"))	
			objVoucher.setMaxGeneration(objRS("max_generation"))	
			objVoucher.setMaxUsage(objRS("max_usage"))	
			objVoucher.setEnableDate(objRS("enable_date"))	
			objVoucher.setExpireDate(objRS("expire_date"))		
			objVoucher.setExcludeProdRule(objRS("exclude_prod_rule"))		
			objVoucher.setObjVoucherCode(objVC)	
			
			Set findExtendedVoucherByCode = objVoucher		
			Set objVC = nothing			
			Set objVoucher = nothing					
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function countVoucherCodeByCampaign(id_campaign, id_user_ref)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		countVoucherCodeByCampaign = 0		
		strSQL = "SELECT count(*) as counter FROM voucher_code WHERE voucher_campaign=?"
		if not isNull(id_user_ref) AND not(id_user_ref = "") then
			strSQL = strSQL & " AND id_user_ref=?"
		end if
		strSQL =  strSQL&";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_campaign)
		if not isNull(id_user_ref) AND not(id_user_ref = "") then			
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user_ref)
		end if
		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then								
			countVoucherCodeByCampaign = objRS("counter")					
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function countVoucherCodeByCode(voucher_code)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		countVoucherCodeByCode = 0		
		strSQL = "SELECT count(*) as counter FROM voucher_code WHERE code=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,voucher_code)
		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then								
			countVoucherCodeByCode = objRS("counter")					
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function generateVoucherCode(id_campaign, id_user_ref, objConn)
		on error resume next
		generateVoucherCode = ""

		Set objVC = findCampaignByID(id_campaign)		
		v_type = objVC.getVoucherType()
		mag_gen = objVC.getMaxGeneration()
		generation_counter = countVoucherCodeByCampaign(id_campaign, id_user_ref)

		if(Clng(generation_counter)<Clng(mag_gen) OR Clng(mag_gen)=-1)then
			Set objGUID = new GUIDClass
			new_code = objGUID.CreateVoucherCodeGUID()
						
			Do While (Cint(countVoucherCodeByCode(new_code)) > 0)
				new_code = objGUID.CreateVoucherCodeGUID()
			Loop			
			
			Set objGUID = nothing
			
			Select Case CInt(v_type)
			Case 0,1,2,3
				generateVoucherCode = insertVoucherCode(new_code, id_campaign, now(), 0, id_user_ref, objConn)				
			Case 4			
				if (not(isNull(id_user_ref)) AND id_user_ref<>"") then
					generateVoucherCode = insertVoucherCode(new_code, id_campaign, now(), 0, id_user_ref, objConn)
				end if
			Case else
			End Select
		end if		
		
		Set objVC = nothing

 		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function generateVoucherCodeByCampaignLabel(campaign_label, id_user_ref, objConn)
		on error resume next
		generateVoucherCodeByCampaignLabel = ""
		
		Set objVC = findCampaignByLabel(campaign_label)
		id_campaign=	objVC.getID()	
		v_type = objVC.getVoucherType()
		mag_gen = objVC.getMaxGeneration()
		generation_counter = countVoucherCodeByCampaign(id_campaign, id_user_ref)

		if(Clng(generation_counter)<Clng(mag_gen) OR Clng(mag_gen)=-1)then
			Set objGUID = new GUIDClass
			new_code = objGUID.CreateVoucherCodeGUID()
			
			Do While (Cint(countVoucherCodeByCode(new_code)) > 0)
				new_code = objGUID.CreateVoucherCodeGUID()
			Loop
			
			Set objGUID = nothing
			
			Select Case CInt(v_type)
			Case 0,1,2,3
				generateVoucherCodeByCampaignLabel = insertVoucherCode(new_code, id_campaign, now(), 0, id_user_ref, objConn)				
			Case 4			
				if (not(isNull(id_user_ref)) AND id_user_ref<>"") then
					generateVoucherCodeByCampaignLabel = insertVoucherCode(new_code, id_campaign, now(), 0, id_user_ref, objConn)
				end if
			Case else
			End Select
		end if		
		
		Set objVC = nothing

 		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function insertVoucherCode(code, voucher_campaign, insert_date, usage_counter, id_user_ref, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		insertVoucherCode=""
		
		strSQL = "INSERT INTO voucher_code(code, voucher_campaign, insert_date, usage_counter, id_user_ref) VALUES("
		strSQL = strSQL & "?,?,?,?,"
		if(isNull(id_user_ref) OR id_user_ref = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if		
		strSQL = strSQL & ");"
						
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,code)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,voucher_campaign)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,insert_date)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,usage_counter)
		if not isNull(id_user_ref) AND not(id_user_ref = "") then			
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user_ref)
		end if
		objCommand.Execute()
		Set objCommand = Nothing

		insertVoucherCode = code

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyVoucherCode(id, code, voucher_campaign, insert_date, usage_counter, id_user_ref, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE voucher_code SET "
		strSQL = strSQL & "`code`=?,"	
		strSQL = strSQL & "voucher_campaign=?,"	
		strSQL = strSQL & "insert_date=?,"		
		strSQL = strSQL & "`usage_counter`=?,"
		if(isNull(id_user_ref) OR id_user_ref = "") then
			strSQL = strSQL & "id_user_ref=NULL"
		else
			strSQL = strSQL & "id_user_ref=?"
		end if
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,code)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,voucher_campaign)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,insert_date)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,usage_counter)
		if not isNull(id_user_ref) AND not(id_user_ref = "") then			
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user_ref)
		end if
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
		
	Public Sub deleteVoucherCodeNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM voucher_code WHERE id=?;"
		
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
		
	Public Sub deleteVoucherCode(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM voucher_code WHERE id=?;"
		
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
		
	Public Sub deleteVoucherCodeByCampaignNoTransaction(id_campaign)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM voucher_code WHERE voucher_campaign=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_campaign)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteVoucherCodeByCampaign(id_campaign, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM voucher_code WHERE voucher_campaign=?;"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_campaign)
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
	
	Public Function convertDate(dateToConvert)
		Dim DD, MM, YY, HH, MIN, SS
		
		convertDate = null
		
		DD = DatePart("d", dateToConvert)
		MM = DatePart("m", dateToConvert)
		YY = DatePart("yyyy", dateToConvert)
		HH = DatePart("h", dateToConvert)
		MIN = DatePart("n", dateToConvert)
		SS = DatePart("s", dateToConvert)
		
		convertDate = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS		
	End Function
End Class


'******************************* CREO UNA INNER CLASS PER GENERARE UNA COLLECTION DA UTILIZZARE NELL'ELABORAZIONE DELLE VOUCHER CAMPAIGN
Class VoucherCodeClass
	Private id
	Private voucher_code
	Private voucher_campaign
	Private insert_date
	Private usage_counter
	Private id_user_ref
	
	Private id_order
	Private valore

	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub

	Public Function getVoucherCode()
		getVoucherCode = voucher_code
	End Function
	
	Public Sub setVoucherCode(strCode)
		voucher_code = strCode
	End Sub

	Public Function getVoucherCampaign()
		getVoucherCampaign = voucher_campaign
	End Function
	
	Public Sub setVoucherCampaign(strCampaign)
		voucher_campaign = strCampaign
	End Sub

	Public Function getInsertDate()
		getInsertDate = insert_date
	End Function
	
	Public Sub setInsertDate(strInsertDate)
		insert_date = strInsertDate
	End Sub

	Public Function getUsageCounter()
		getUsageCounter = usage_counter
	End Function
	
	Public Sub setUsageCounter(strUsageCounter)
		usage_counter = strUsageCounter
	End Sub
	
	Public Function getIdUserRef()
		getIdUserRef = id_user_ref
	End Function
	
	Public Sub setIdUserRef(strIdUserRef)
		id_user_ref = strIdUserRef
	End Sub


	Public Function getOrderID()
		getOrderID = id_order
	End Function
	
	Public Sub setOrderID(strOrderID)
		id_order = strOrderID
	End Sub		

	Public Function getValore()
		getValore = valore
	End Function
	
	Public Sub setValore(strValore)
		valore = strValore
	End Sub
End Class
%>