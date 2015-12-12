<%

Class MarginDiscountClass
	Private id
	Private margin
	Private discount
	Private applyProdDiscount
	Private applyUserDiscount
	
	
	Public Function getID()
		getID = id
	End Function
				
	Public Sub setID(numID)
		id = numID
	End Sub
	
	Public Function getMargin()
		getMargin = Cdbl(margin)
	End Function
	
	Public Sub setMargin(dblMargin)
		margin = dblMargin
	End Sub
	
	Public Function getDiscount()
		getDiscount = Cdbl(discount)
	End Function
	
	Public Sub setDiscount(dblDiscount)
		discount = dblDiscount
	End Sub
	
	Public Function isApplyProdDiscount()
		isApplyProdDiscount = applyProdDiscount
	End Function
	
	Public Sub setApplyProdDiscount(bolApplyProdDiscount)
		applyProdDiscount = bolApplyProdDiscount
	End Sub
	
	Public Function isApplyUserDiscount()
		isApplyUserDiscount = applyUserDiscount
	End Function
	
	Public Sub setApplyUserDiscount(bolApplyUserDiscount)
		applyUserDiscount = bolApplyUserDiscount
	End Sub



'*********************************** METODI MARGIN DISCOUNT *********************** 			
		
	Public Function insertMarginDiscount(dblMargin, dblDiscount, bolApplyProdDiscount, bolApplyUserDiscount, objConn)
		on error resume next
		insertMarginDiscount = -1
		
		Dim strSQL, strSQLSelect, objRS
		
		strSQL = "INSERT INTO margin_discount(margin, discount, apply_prod_discount, apply_user_discount) VALUES("
		strSQL = strSQL & "?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(dblMargin))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(dblDiscount))
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolApplyProdDiscount)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolApplyUserDiscount)
		objCommand.Execute()
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(margin_discount.id) as id FROM margin_discount")
		
		if not (objRS.EOF) then
			'objRS.MoveFirst()
			insertMarginDiscount = objRS("id")	
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
		
	Public Sub modifyMarginDiscount(id, dblMargin, dblDiscount, bolApplyProdDiscount, bolApplyUserDiscount, objConn)
		on error resume next
		Dim strSQL, objRS
		
		strSQL = "UPDATE margin_discount SET "
		strSQL = strSQL & "margin=?,"
		strSQL = strSQL & "discount=?,"
		strSQL = strSQL & "apply_prod_discount=?,"
		strSQL = strSQL & "apply_user_discount=?"
		strSQL = strSQL & " WHERE id=?;" 

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(dblMargin))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(dblDiscount))
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolApplyProdDiscount)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolApplyUserDiscount)
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
		
	Public Sub deleteMarginDiscount(id)
		on error resume next
		Dim objDB, strSQLDel, objConn, strSQLDelMarginGroup
		strSQLDel = "DELETE FROM margin_discount WHERE id=?;" 
		strSQLDelMarginGroup = "DELETE FROM usr_group_x_margin_disc WHERE id_marg_disc=?;"

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
		
	Public Function getListaMarginDiscount()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objMarginDiscount
		getListaMarginDiscount = null  
		strSQL = "SELECT * FROM margin_discount;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Set objRS = objConn.Execute(strSQL)
		
		if not(objRS.EOF) then		   
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
			do while not objRS.EOF
				Set objMarginDiscount = new MarginDiscountClass
				strID = objRS("id")
								
				objMarginDiscount.setID(strID)    
				objMarginDiscount.setMargin(objRS("margin"))
				objMarginDiscount.setDiscount(objRS("discount"))
				objMarginDiscount.setApplyProdDiscount(objRS("apply_prod_discount"))
				objMarginDiscount.setApplyUserDiscount(objRS("apply_user_discount"))
									
				objDict.add strID, objMarginDiscount
				Set objMarginDiscount = nothing
				objRS.moveNext()
			loop
			
			Set getListaMarginDiscount = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
				
	Public Function findMarginDiscountByID(id)
		on error resume next
		
		findMarginDiscountByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM margin_discount WHERE id=?;"

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
			Dim objMarginDiscount
			Set objMarginDiscount = new MarginDiscountClass	
			
			objMarginDiscount.setID(objRS("id"))    
			objMarginDiscount.setMargin(objRS("margin"))
			objMarginDiscount.setDiscount(objRS("discount"))
			objMarginDiscount.setApplyProdDiscount(objRS("apply_prod_discount"))
			objMarginDiscount.setApplyUserDiscount(objRS("apply_user_discount"))
			
			Set findMarginDiscountByID = objMarginDiscount
			Set objMarginDiscount = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	

	Public Function getAmount(dblAmount,dblMargin,dblDiscount,applyProdDisc,applyUserDisc,prodDiscount,userDiscount)
		'*************** gestione della logica di calcolo dei margini/sconti
		'*************** in associazione con lo sconto per il cliente e lo sconto per il prodotto

		'*************** verifico se devono essere applicati gli sconti prodotto e cliente, e li aggiungo allo sconto esistente
		if(applyProdDisc=1)then
			dblDiscount = dblDiscount+prodDiscount
		end if
		if(applyUserDisc=1)then
			dblDiscount = dblDiscount+userDiscount
		end if

		dblMargin = dblMargin-dblDiscount

		getAmount = dblAmount + (dblAmount / 100 * dblMargin)	
	End Function	


	Public Function getMarginAmount(dblAmount,dblMargin)
		getMarginAmount = (dblAmount / 100 * dblMargin)	
	End Function
	

	Public Function getDiscountAmount(dblAmount,dblDiscount,applyProdDisc,applyUserDisc,prodDiscount,userDiscount)
		if(applyProdDisc=1)then
			dblDiscount = dblDiscount+prodDiscount
		end if
		if(applyUserDisc=1)then
			dblDiscount = dblDiscount+userDiscount
		end if

		'dblDiscount = -dblDiscount

		getDiscountAmount = (dblAmount / 100 * dblDiscount)	
	End Function	
	

	Public Function getDiscountPercentual(dblDiscount,applyProdDisc,applyUserDisc,prodDiscount,userDiscount)
		
		if(applyProdDisc=1)then
			dblDiscount = dblDiscount+prodDiscount
		end if
		if(applyUserDisc=1)then
			dblDiscount = dblDiscount+userDiscount
		end if

		getDiscountPercentual = dblDiscount	
	End Function
	

	Public Sub insertMarginDiscountXUserGroup(id_margin_discount, id_user_group, objConn)
		on error resume next
		Dim strSQL, objRS
		
		strSQL = "INSERT INTO usr_group_x_margin_disc(id_marg_disc, id_user_group) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_margin_discount)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user_group)
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
	
		
	Public Sub deleteMarginDiscountXUserGroup(id_user_group, objConn)
		on error resume next
		Dim strSQL, objRS 
		strSQL = "DELETE FROM usr_group_x_margin_disc WHERE id_user_group=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user_group)
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

	'public Sub toString()
		'response.write ()
	'end Sub
End Class
%>