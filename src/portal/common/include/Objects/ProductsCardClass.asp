<%

Class ProductsCardClass
	
	Private idCarrello
	Private idProdotto
	Private counterProd
	Private qtaProdotto
	Private prod_type
	
	
	'***************** INIZIO: METODI PER GESTIRE I SINGOLI PRODOTTI PER CARRELLO
	Public Function getIDCarrello()
		getIDCarrello = idCarrello
	End Function	

	Public Sub setIDCarrello(idCarr)
		idCarrello = idCarr
	End Sub		
	
	Public Function getIDProd()
		getIDProd = idProdotto
	End Function	

	Public Sub setIDProd(idProd)
		idProdotto = idProd
	End Sub				
	
	Public Function getCounterProd()
		getCounterProd = counterProd
	End Function	
	
	Public Sub setCounterProd(numCounterProd)
		counterProd = numCounterProd
	End Sub	
	
	Public Function getQtaProd()
		getQtaProd = qtaProdotto
	End Function	

	Public Sub setQtaProd(qtaProd)
		qtaProdotto = qtaProd
	End Sub	

	Public Function getProdType()
		getProdType = prod_type
	End Function
	
	Public Sub setProdType(intProdType)
		prod_type = intProdType
	End Sub		
		
	
	Public Function retrieveListaProdotti(id_carrello)
		on error resume next
		retrieveListaProdotti = null
		
		Dim objDB, strSQLSelect, objRS, objConn 
		dim numIDProd, objProdTmp, objListaProd
		
		strSQLSelect = "SELECT * FROM prodotti_x_carrello WHERE id_carrello=? ORDER BY id_prodotto,counter_prod;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_carrello)
		Set objRS = objCommand.Execute()

		if not objRS.EOF then
			Set objListaProd = Server.CreateObject("Scripting.Dictionary") 
			do while not objRS.EOF
				Set objProdTmp = New ProductsCardClass
				numIDProd = objRS("id_prodotto")
				counterProd = objRS("counter_prod")
				objProdTmp.setIDCarrello(objRS("id_carrello"))
				objProdTmp.setIDProd(objRS("id_prodotto"))
				objProdTmp.setCounterProd(counterProd)
				objProdTmp.setQtaProd(objRS("qta_prod"))
				objProdTmp.setProdType(objRS("prod_type"))
				objListaProd.Add numIDProd&"|"&counterProd, objProdTmp 
				Set objProdTmp = nothing
				objRS.moveNext()
			loop
			
			Set retrieveListaProdotti = objListaProd
			Set objListaProd = Nothing	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function 

	Public Sub addItem(numIDCarrello, numIDProd, numCounterProd, qtaProd, prod_type, resetQta, objConn)
		on error resume next
		Dim strSQLInsert, strSQLUpdate, objCarrello, objCurrCarrello		
		Set objCarrello = New CardClass
		Set objCurrCarrello = objCarrello.getCarrelloByIDCarello(numIDCarrello)
		
		if not(isNull(objCurrCarrello)) AND not(isEmpty(objCurrCarrello)) then
			Dim objProdPerCarr
			Set objProdPerCarr = New ProductsCardClass

			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1

			if(objProdPerCarr.existItem(numIDCarrello, numIDProd, numCounterProd)) then
				Dim tmpQta, tmpTotale, tmpCarrTotProd, objTmpProdCarr
				
				'**** gestisco la quantit�, se resetQta � false aggiungo alla quantit� gi� presente, altrimenti imposto alla nuova quantit� passata				
				if (resetQta) then
					tmpQta = qtaProd
				else
					Set objTmpProdCarr = objProdPerCarr.getItem(numIDCarrello, numIDProd, numCounterProd)
					tmpQta = qtaProd + objTmpProdCarr.getQtaProd()		
					Set objTmpProdCarr = nothing				
				end if
				
				strSQLUpdate = "UPDATE prodotti_x_carrello SET "
				strSQLUpdate = strSQLUpdate & "qta_prod=?"
				strSQLUpdate = strSQLUpdate & " WHERE id_prodotto=?"
				strSQLUpdate = strSQLUpdate & " AND id_carrello=?"
				strSQLUpdate = strSQLUpdate & " AND counter_prod=?;"

				objCommand.CommandText = strSQLUpdate
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tmpQta)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDProd)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDCarrello)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numCounterProd)
				objCommand.Execute()
			else				
				strSQLInsert = "INSERT INTO prodotti_x_carrello(id_carrello, id_prodotto, counter_prod, qta_prod, prod_type) VALUES("
				strSQLInsert = strSQLInsert & "?,?,?,?,?);"
				
				objCommand.CommandText = strSQLInsert
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDCarrello)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDProd)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numCounterProd)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qtaProd)
				objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
				objCommand.Execute()
			end if	
			Set objCommand = Nothing
		
			if objConn.Errors.Count > 0 then
				objConn.RollBackTrans
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if		
		
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=019")
		end if	
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Sub	
	 
	Public Sub addItemNoTransaction(numIDCarrello, numIDProd, numCounterProd, qtaProd, prod_type, resetQta)
		on error resume next
		Dim objDB, strSQLInsert, strSQLUpdate, objCarrello, objCurrCarrello		
		Set objCarrello = New CardClass
		Set objCurrCarrello = objCarrello.getCarrelloByIDCarello(numIDCarrello)
		
		if not(isNull(objCurrCarrello)) AND not(isEmpty(objCurrCarrello)) then
			Dim objProdPerCarr
			Set objProdPerCarr = New ProductsCardClass			
			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()	

			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1

			objConn.BeginTrans			
			
			if(objProdPerCarr.existItem(numIDCarrello, numIDProd, numCounterProd)) then
				Dim tmpQta, tmpTotale, tmpCarrTotProd, objTmpProdCarr
				
				'**** gestisco la quantit�, se resetQta � false aggiungo alla quantit� gi� presente, altrimenti imposto alla nuova quantit� passata				
				if (resetQta) then
					tmpQta = qtaProd
				else
					Set objTmpProdCarr = objProdPerCarr.getItem(numIDCarrello, numIDProd, numCounterProd)
					tmpQta = qtaProd + objTmpProdCarr.getQtaProd()		
					Set objTmpProdCarr = nothing				
				end if
								
				strSQLUpdate = "UPDATE prodotti_x_carrello SET "				
				strSQLUpdate = strSQLUpdate & "qta_prod=?"
				strSQLUpdate = strSQLUpdate & " WHERE id_prodotto=?"
				strSQLUpdate = strSQLUpdate & " AND id_carrello=?"
				strSQLUpdate = strSQLUpdate & " AND counter_prod=?;"
				
				objCommand.CommandText = strSQLUpdate
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tmpQta)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDProd)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDCarrello)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numCounterProd)
				objCommand.Execute()
			else				
				strSQLInsert = "INSERT INTO prodotti_x_carrello(id_carrello, id_prodotto, counter_prod, qta_prod, prod_type) VALUES("
				strSQLInsert = strSQLInsert & "?,?,?,?);"				

				objCommand.CommandText = strSQLInsert
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDCarrello)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDProd)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numCounterProd)
				objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qtaProd)
				objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
				objCommand.Execute()
			end if				
			Set objCommand = Nothing

			if objConn.Errors.Count = 0 then
				objConn.CommitTrans
			end If
		
			if objConn.Errors.Count > 0 then
				objConn.RollBackTrans
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			
			Set objDB = nothing		
		
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=019")
		end if	
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Sub	

	Public Sub delItem(numIDCarrello, numIDProd, numCounterProd, objConn)
		on error resume next
		Dim objDB, strSQLDel, objCarrello, objCurrCarrello
		
		Set objCarrello = New CardClass
		Set objCurrCarrello = objCarrello.getCarrelloByIDCarello(numIDCarrello)
		
		if not(isNull(objCurrCarrello)) AND not(isEmpty(objCurrCarrello)) then
			strSQLDel = "DELETE FROM prodotti_x_carrello WHERE id_prodotto=? AND id_carrello=? AND counter_prod=?;"

			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQLDel
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDProd)
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDCarrello)
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numCounterProd)
			objCommand.Execute()			
			Set objCommand = Nothing
		
			if objConn.Errors.Count > 0 then
				objConn.RollBackTrans
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
									
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=019")
		end if				
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Sub
	
	Public Sub delItemNoTransaction(numIDCarrello, numIDProd, numCounterProd)
		on error resume next
		Dim objDB, strSQLDel, objCarrello, objCurrCarrello
		
		Set objCarrello = New CardClass
		Set objCurrCarrello = objCarrello.getCarrelloByIDCarello(numIDCarrello)
		
		if not(isNull(objCurrCarrello)) AND not(isEmpty(objCurrCarrello)) then
			strSQLDel = "DELETE FROM prodotti_x_carrello WHERE id_prodotto=? AND id_carrello=? AND counter_prod=?;"
				
			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()	

			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQLDel
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDProd)
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDCarrello)
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numCounterProd)
			
			objConn.BeginTrans			

			objCommand.Execute()					
			if objConn.Errors.Count = 0 then
				objConn.CommitTrans
			end If
		
			if objConn.Errors.Count > 0 then
				objConn.RollBackTrans
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			
			Set objCommand = Nothing
			Set objDB = nothing								
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=019")
		end if				
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Sub
				
	Public Function existItem(numIDCarrello, numIDProd, numCounterProd)
		on error resume next
		Dim objDB, strSQLSelect, objRS, objConn
		
		existItem = false
		
		strSQLSelect = "SELECT * FROM prodotti_x_carrello WHERE id_carrello=? AND id_prodotto=?"
		if not(isNull(numCounterProd))then
			strSQLSelect = strSQLSelect & " AND counter_prod=?"
		end if
		strSQL = strSQL & ";"
						
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDCarrello)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDProd)
		if not(isNull(numCounterProd))then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numCounterProd)
		end if
		Set objRS = objCommand.Execute()
		
		if objRS.EOF then
			existItem = false	
		else
			existItem = true
		end if
		
		objRS.Close
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function		
		
	Public Function getItem(numIDCarrello, numIDProd, numCounterProd)
		on error resume next
		Dim objDB, strSQLSelect, objRS, objConn, objProd
		
		getItem = null
		
		strSQLSelect = "SELECT * FROM prodotti_x_carrello WHERE id_carrello=? AND id_prodotto=? AND counter_prod=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDCarrello)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numCounterProd)
		Set objRS = objCommand.Execute()
		
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=019")
		else
			Set objProd = New ProductsCardClass
			objProd.setIDCarrello(objRS("id_carrello"))
			objProd.setIDProd(objRS("id_prodotto"))
			objProd.setCounterProd(objRS("counter_prod"))
			objProd.setQtaProd(objRS("qta_prod"))
			objProd.setProdType(objRS("prod_type"))
			Set getItem = objProd
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function		
		
	Public Function getListItem(numIDCarrello, numIDProd)
		on error resume next
		Dim objDB, strSQLSelect, objRS, objConn, objProd
		
		getListItem = null
		
		strSQLSelect = "SELECT * FROM prodotti_x_carrello WHERE id_carrello=? AND id_prodotto=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDCarrello)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDProd)
		Set objRS = objCommand.Execute()
		
		if not objRS.EOF then
			Set objListaProd = Server.CreateObject("Scripting.Dictionary") 
			do while not objRS.EOF
				Set objProdTmp = New ProductsCardClass
				numIDProd = objRS("id_prodotto")
				counterProd = objRS("counter_prod")
				objProdTmp.setIDCarrello(objRS("id_carrello"))
				objProdTmp.setIDProd(objRS("id_prodotto"))
				objProdTmp.setCounterProd(counterProd)
				objProdTmp.setQtaProd(objRS("qta_prod"))
				objProd.setProdType(objRS("prod_type"))
				objListaProd.Add numIDProd&"|"&counterProd, objProdTmp 
				Set objProdTmp = nothing
				objRS.moveNext()
			loop
			
			Set getListItem = objListaProd
			Set objListaProd = Nothing	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function	
			
	Public Function getMaxItemCounterProd(numIDCarrello, numIDProd)
		on error resume next
		Dim objDB, strSQLSelect, objRS, objConn, objProd
		
		getMaxItemCounterProd = -1
		
		strSQLSelect = "SELECT counter_prod FROM prodotti_x_carrello WHERE id_carrello=? AND id_prodotto=? ORDER BY counter_prod DESC LIMIT 1;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDCarrello)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDProd)
		Set objRS = objCommand.Execute()
		
		if objRS.EOF then
			getMaxItemCounterProd = -1
		else
			getMaxItemCounterProd = objRS("counter_prod") 
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function
	
	public Sub toString()
		response.write ("ok")
	end Sub
End Class
%>