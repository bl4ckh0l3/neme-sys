<%

Class UserGroupClass
	Private id
	Private shortDesc
	Private longDesc
	Private defaultGroup
	Private taxsGroup
	
	
	Public Function getID()
		getID = id
	End Function
				
	Public Sub setID(numID)
		id = numID
	End Sub
		
	Public Function getShortDesc()
		getShortDesc = shortDesc
	End Function
		
	Public Sub setShortDesc(strShortDesc)
		shortDesc = strShortDesc
	End Sub
		
	Public Function getLongDesc()
		getLongDesc = longDesc
	End Function
		
	Public Sub setLongDesc(strLongDesc)
		longDesc = strLongDesc
	End Sub	
		
	Public Function isDefault()
		isDefault = defaultGroup
	End Function
		
	Public Sub setDefault(bolDefault)
		defaultGroup = bolDefault
	End Sub		
	
	Public Function getTaxGroup()
		getTaxGroup = taxsGroup
	End Function
	
	Public Sub setTaxGroup(strTaxGroup)
		taxsGroup = strTaxGroup
	End Sub
	
	
	Public Function getTaxGroupObj(iTaxGroup)
		getTaxGroupObj = null
		if (not(isNull(iTaxGroup)) AND iTaxGroup<>"")then
			On Error Resume Next
			Set objTG = New TaxsGroupClass
			Set getTaxGroupObj = objTG.getGroupByID(iTaxGroup)			
			Set objTG = nothing
			if(Err.number <> 0) then
				Set getTaxGroupObj = null
			end if
		end if
	End Function
	


'*********************************** METODI USERGROUP  *********************** 		
		
	Public Function insertUserGroup(strShortDesc, strLongDesc, bolDefaultGroup, strTaxGroup)
		on error resume next
		insertUserGroup = -1
		
		Dim objDB, strSQL, strSQLSelect, objRS, objConn
		
		strSQL = "INSERT INTO user_group(short_desc, long_desc,default_group, taxs_group) VALUES("
		strSQL = strSQL & "?,?,?,"
		if(isNull(strTaxGroup) OR strTaxGroup = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ");"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strShortDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strLongDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolDefaultGroup)
		if not isNull(strTaxGroup) AND not(strTaxGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strTaxGroup)
		end if
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(user_group.id) as id FROM user_group")
		if not (objRS.EOF) then
			insertUserGroup = objRS("id")	
		end if	
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function	
		
	Public Sub modifyUserGroup(id, strShortDesc, strLongDesc, bolDefaultGroup,strTaxGroup)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE user_group SET "
		strSQL = strSQL & "short_desc=?,"
		strSQL = strSQL & "long_desc=?,"
		strSQL = strSQL & "default_group=?,"
		if(isNull(strTaxGroup) OR strTaxGroup = "") then
			strSQL = strSQL & "taxs_group=NULL"
		else
			strSQL = strSQL & "taxs_group=?"			
		end if
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strShortDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strLongDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolDefaultGroup)
		if not isNull(strTaxGroup) AND not(strTaxGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strTaxGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub				
		
	Public Sub deleteUserGroup(id)
		on error resume next
		Dim objDB, strSQLDel, objConn
		strSQLDel = "DELETE FROM user_group WHERE id=?;"
		strSQLDelMarginGroup = "DELETE FROM usr_group_x_margin_disc WHERE id_user_group=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		Dim objCommand, objCommand2
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQLDel
		objCommand2.CommandText = strSQLDelMarginGroup
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)

		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand2.Execute()
		end if
		objCommand.Execute()
		
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
			
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
		
	Public Function getListaUserGroup()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objUserGroup
		getListaUserGroup = null  
		strSQL = "SELECT * FROM user_group ORDER BY short_desc ASC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Set objRS = objConn.Execute(strSQL)
		
		if not(objRS.EOF) then		   
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
			do while not objRS.EOF
				Set objUserGroup = new UserGroupClass
				strID = objRS("id")
								
				objUserGroup.setID(strID)    
				objUserGroup.setShortDesc(objRS("short_desc"))
				objUserGroup.setLongDesc(objRS("long_desc"))
				objUserGroup.setDefault(objRS("default_group"))
				objUserGroup.setTaxGroup(objRS("taxs_group"))	
									
				objDict.add strID, objUserGroup
				Set objUserGroup = nothing
				objRS.moveNext()
			loop
			
			Set getListaUserGroup = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
				
	Public Function findUserGroupByID(id)
		on error resume next
		
		findUserGroupByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM user_group WHERE id=?;"

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
			Dim objUserGroup			
			Set objUserGroup = new UserGroupClass			
			objUserGroup.setID(objRS("id"))    
			objUserGroup.setShortDesc(objRS("short_desc"))
			objUserGroup.setLongDesc(objRS("long_desc"))
			objUserGroup.setDefault(objRS("default_group"))
			objUserGroup.setTaxGroup(objRS("taxs_group"))	
						
			Set findUserGroupByID = objUserGroup
			Set objUserGroup = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
				
	Public Function findUserGroupDefault()
		on error resume next
		
		findUserGroupDefault = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM user_group WHERE default_group=1;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		
		if not(objRS.EOF) then
			Dim objUserGroup
			
			Set objUserGroup = new UserGroupClass	
			
			objUserGroup.setID(objRS("id"))    
			objUserGroup.setShortDesc(objRS("short_desc"))
			objUserGroup.setLongDesc(objRS("long_desc"))
			objUserGroup.setDefault(objRS("default_group"))
			objUserGroup.setTaxGroup(objRS("taxs_group"))	
						
			Set findUserGroupDefault = objUserGroup
			Set objUserGroup = Nothing
		end if		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getMaxIDUserGroup()
		on error resume next
		
		getMaxIDUserGroup = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT MAX(id) AS id FROM user_group;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDUserGroup = objRS("id")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function countUserGroup()
		on error resume next
		
		countUserGroup = 0
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT count(*) AS counter FROM user_group;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		
		if not(objRS.EOF) then
			countUserGroup = objRS("counter")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function getUserGroupXMarginDiscount(strIDM)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, strIDGroup, strIDMargin
		getUserGroupXMarginDiscount = null 
		strSQL = "SELECT user_group.* FROM user_group INNER JOIN usr_group_x_margin_disc ON user_group.id = usr_group_x_margin_disc.id_user_group"
		if not(isNull(strIDM))then
			strSQL = strSQL &" WHERE usr_group_x_margin_disc.id_marg_disc=?"
		end if
		strSQL = strSQL &";"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if not(isNull(strIDM))then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strIDM)
		end if
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then		   
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
			do while not objRS.EOF
				Set objUserGroup = new UserGroupClass
				strID = objRS("id")
								
				objUserGroup.setID(strID)    
				objUserGroup.setShortDesc(objRS("short_desc"))
				objUserGroup.setLongDesc(objRS("long_desc"))
				objUserGroup.setDefault(objRS("default_group"))
				objUserGroup.setTaxGroup(objRS("taxs_group"))	
									
				objDict.add strID, objUserGroup
				Set objUserGroup = nothing
				objRS.moveNext()
			loop
			
			Set getUserGroupXMarginDiscount = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function

	Public Function getMarginDiscountXUserGroup(id_user_group)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		getMarginDiscountXUserGroup = null		
		strSQL = "SELECT margin_discount.* FROM margin_discount INNER JOIN usr_group_x_margin_disc ON margin_discount.id = usr_group_x_margin_disc.id_marg_disc WHERE usr_group_x_margin_disc.id_user_group=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user_group)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			Dim objMarginDiscount		
			Set objMarginDiscount = new MarginDiscountClass	
			
			objMarginDiscount.setID(objRS("id"))    
			objMarginDiscount.setMargin(objRS("margin"))
			objMarginDiscount.setDiscount(objRS("discount"))
			objMarginDiscount.setApplyProdDiscount(objRS("apply_prod_discount"))
			objMarginDiscount.setApplyUserDiscount(objRS("apply_user_discount"))
						
			Set getMarginDiscountXUserGroup = objMarginDiscount
			Set objMarginDiscount = Nothing
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