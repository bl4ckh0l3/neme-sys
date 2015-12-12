<%

Class Products4OrderClass
	Private idOrdine
	Private idProdotto
	Private counterProd
	Private nomeProdotto
	Private qtaProdotto
	Private totale
	Private tax
	Private descTax
	Private prod_type
	
	Public Function getIDOrdine()
		getIDOrdine = idOrdine
	End Function	
	
	Public Function getIDProdotto()
		getIDProdotto = idProdotto
	End Function
	
	Public Function getCounterProd()
		getCounterProd = counterProd
	End Function	
	
	Public Function getNomeProdotto()
		getNomeProdotto = nomeProdotto
	End Function
		
	Public Function getQtaProdotto()
		getQtaProdotto = qtaProdotto
	End Function	
	
	Public Function getTotale()
		getTotale = Cdbl(totale)
	End Function	
	
	Public Function getTax()
		getTax = Cdbl(tax)
	End Function	
	
	Public Function getDescTax()
		getDescTax = descTax
	End Function

	Public Function getProdType()
		getProdType = prod_type
	End Function
		
		
	Public Sub setIDOrdine(numIDOrdine)
		idOrdine = numIDOrdine
	End Sub
			
	Public Sub setIDProdotto(numIDProdotto)
		idProdotto = numIDProdotto
	End Sub
	
	Public Sub setCounterProd(numCounterProd)
		counterProd = numCounterProd
	End Sub	
			
	Public Sub setNomeProdotto(strNomeProdotto)
		nomeProdotto = strNomeProdotto
	End Sub
			
	Public Sub setQtaProdotto(numQtaProdotto)
		qtaProdotto = numQtaProdotto
	End Sub
			
	Public Sub setTotale(numTotale)
		totale = numTotale
	End Sub
			
	Public Sub setTax(numTax)
		tax = numTax
	End Sub
			
	Public Sub setDescTax(strDescTax)
		descTax = strDescTax
	End Sub
	
	Public Sub setProdType(intProdType)
		prod_type = intProdType
	End Sub
		
	
	'*************************** METODI PRODOTTI PER ORDINE ***********************
	Public Sub insertProdottiXOrdineNoTransaction(id_ordine, id_prodotto, counter_prod, nome_prodotto, qta, totale, tax, desc_tax, prod_type)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		
		strSQL = "INSERT INTO prodotti_x_ordine(id_ordine, id_prodotto, counter_prod, nome_prodotto, qta, totale, tax, desc_tax, prod_type) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,?);"
						
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,nome_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qta)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(totale))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(tax))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,desc_tax)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyProdottiXOrdineNoTransaction(id, id_ordine, id_prodotto, counter_prod, nome_prodotto, qta, totale, tax, desc_tax, prod_type)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		strSQL = "UPDATE prodotti_x_ordine SET "
		strSQL = strSQL & "id_ordine=?,"
		strSQL = strSQL & "id_prodotto=?,"
		strSQL = strSQL & "counter_prod=?,"
		strSQL = strSQL & "nome_prodotto=?,"
		strSQL = strSQL & "qta=?,"
		strSQL = strSQL & "totale=?,"
		strSQL = strSQL & "tax=?,"
		strSQL = strSQL & "desc_tax=?,"
		strSQL = strSQL & "prod_type=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,nome_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qta)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(totale))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(tax))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,desc_tax)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
	
	Public Sub insertProdottiXOrdine(id_ordine, id_prodotto, counter_prod, nome_prodotto, qta, totale, tax, desc_tax, prod_type, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO prodotti_x_ordine(id_ordine, id_prodotto, counter_prod, nome_prodotto, qta, totale, tax, desc_tax, prod_type) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,?);"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,nome_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qta)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(totale))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(tax))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,desc_tax)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
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
		
	Public Sub modifyProdottiXOrdine(id, id_ordine, id_prodotto, counter_prod, nome_prodotto, qta, totale, tax, desc_tax, prod_type, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE prodotti_x_ordine SET "
		strSQL = strSQL & "id_ordine=?,"
		strSQL = strSQL & "id_prodotto=?,"
		strSQL = strSQL & "counter_prod=?,"
		strSQL = strSQL & "nome_prodotto=?,"
		strSQL = strSQL & "qta=?,"
		strSQL = strSQL & "totale=?,"
		strSQL = strSQL & "tax=?,"
		strSQL = strSQL & "desc_tax=?,"
		strSQL = strSQL & "prod_type=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,nome_prodotto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,qta)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(totale))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(tax))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,desc_tax)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
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
		
	Public Sub deleteProdottiXOrdineNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn 
		strSQL = "DELETE FROM prodotti_x_ordine WHERE id_ordine=?;"

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
		
	Public Sub deleteProdottiXOrdine(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM prodotti_x_ordine WHERE id_ordine=?;"

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
	
	Public Sub deleteProdottoXOrdineNoTransaction(id_ordine, id_prod)
		on error resume next
		Dim objDB, strSQL, objRS, objConn 
		strSQL = "DELETE FROM prodotti_x_ordine WHERE id_ordine=? AND id_prodotto=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub deleteProdottoXOrdine(id_ordine, id_prod, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM prodotti_x_ordine WHERE id_ordine=? AND id_prodotto=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
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
	
	Public Sub deleteProdottoXOrdProdCountNoTransaction(id_ordine, id_prod, counter_prod)
		on error resume next
		Dim objDB, strSQL, objRS, objConn 
		strSQL = "DELETE FROM prodotti_x_ordine WHERE id_ordine=? AND id_prodotto=? AND counter_prod=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter_prod)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub deleteProdottoXOrdProdCount(id_ordine, id_prod, counter_prod, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM prodotti_x_ordine WHERE id_ordine=? AND id_prodotto=? AND counter_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ordine)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,counter_prod)
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
				
	Public Function getListaProdottiXOrdine(id_ordine)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict		
		getListaProdottiXOrdine = null		
		strSQL = "SELECT * FROM prodotti_x_ordine WHERE id_ordine=? ORDER BY id_prodotto,counter_prod;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ordine)
		Set objRS = objCommand.Execute()		
		
		if objRS.EOF then			
		else
			Dim objProdOrd, idProdTmp		
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objProdOrd = New Products4OrderClass
				idProdTmp = objRS("id_prodotto")
				counterProdTmp = objRS("counter_prod")
				objProdOrd.setIDOrdine(objRS("id_ordine"))
				objProdOrd.setIDProdotto(idProdTmp)
				objProdOrd.setCounterProd(counterProdTmp)
				objProdOrd.setNomeProdotto(objRS("nome_prodotto"))
				objProdOrd.setQtaProdotto(objRS("qta"))
				objProdOrd.setTotale(objRS("totale"))	
				objProdOrd.setTax(objRS("tax"))	
				objProdOrd.setDescTax(objRS("desc_tax"))	
				objProdOrd.setProdType(objRS("prod_type"))		
				objDict.add idProdTmp&"|"&counterProdTmp, objProdOrd	
				Set objProdOrd = nothing
				objRS.moveNext()
			loop
			
			Set getListaProdottiXOrdine = objDict
			Set objDict = nothing			
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
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