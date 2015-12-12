<%
Class OrderClass
	Private idOrdine
	Private idUtente
	Private dtaInserimento
	Private totaleImponibile
	Private totaleTasse
	Private totale
	Private tipoPagam
	Private paymentCommission
	Private pagamEffettuato
	Private statoOrdine
	Private objProdottiOrdine
	Private orderGUID
	Private userNotifiedXDownload
	Private orderNotes
	Private noRegistration
	Private idAdRef
	
	
	Public Function getIDOrdine()
		getIDOrdine = idOrdine
	End Function
	
	Public Function getIDUtente()
		getIDUtente = idUtente
	End Function
			
	Public Function getDtaInserimento()
		getDtaInserimento = dtaInserimento
	End Function
	
	Public Function getTotaleImponibile()
		getTotaleImponibile = Cdbl(totaleImponibile)
	End Function
	
	Public Function getTotaleTasse()
		getTotaleTasse = Cdbl(totaleTasse)
	End Function
	
	Public Function getTotale()
		getTotale = Cdbl(totale)
	End Function
	
	Public Function getTipoPagam()
		getTipoPagam = tipoPagam
	End Function
	
	Public Function getPaymentCommission()
		getPaymentCommission = paymentCommission
	End Function
	
	Public Function getPagamEffettuato()
		getPagamEffettuato = pagamEffettuato
	End Function
	
	Public Function getStatoOrdine()
		getStatoOrdine = statoOrdine
	End Function
	
	Public Function isUserNotifiedXDownload()
		isUserNotifiedXDownload = userNotifiedXDownload
	End Function
	
	Public Function getOrderGUID()
		getOrderGUID = orderGUID
	End Function
	
	Public Function getOrderNotes()
		getOrderNotes = orderNotes
	End Function
	
	Public Function getNoRegistration()
		getNoRegistration = noRegistration
	End Function
	
	Public Function getIdAdRef()
		getIdAdRef = idAdRef
	End Function
	
	
	Public Function getProdottiXOrdine()		
		if(isNull(objProdottiOrdine)) then
			getProdottiXOrdine = null
		else
			Set getProdottiXOrdine = objProdottiOrdine
		end if
	End Function
	
	Public Sub setProdottiXOrdine(objProd)
		if(isNull(objProd)) then
			objProdottiOrdine = null
		else
			Set objProdottiOrdine = objProd
		end if		
	End Sub				
			
	Public Sub setIDOrdine(numIDOrdine)
		idOrdine = numIDOrdine
	End Sub
			
	Public Sub setIDUtente(numIDUtente)
		idUtente = numIDUtente
	End Sub
			
	Public Sub setDtaInserimento(dtaIns)
		dtaInserimento = dtaIns
	End Sub
	
	Public Sub setTotaleImponibile(numTotaleImponibile)
		totaleImponibile = numTotaleImponibile
	End Sub
	
	Public Sub setTotaleTasse(numTotaleTasse)
		totaleTasse = numTotaleTasse
	End Sub
	
	Public Sub setTotale(numTotale)
		totale = numTotale
	End Sub
	
	Public Sub setTipoPagam(strTipoPagam)
		tipoPagam = strTipoPagam
	End Sub
	
	Public Sub setPaymentCommission(dblPaymentCommission)
		paymentCommission = dblPaymentCommission
	End Sub
	
	Public Sub setPagamEffettuato(strPagamEffettuato)
		pagamEffettuato = strPagamEffettuato
	End Sub
	
	Public Sub setStatoOrdine(bolStatoOrd)
		statoOrdine = bolStatoOrd
	End Sub
	
	Public Sub setOrderGUID(strOrderGUID)
		orderGUID = strOrderGUID
	End Sub
	
	Public Sub setUserNotifiedXDownload(strUserNotifiedXDownload)
		userNotifiedXDownload = strUserNotifiedXDownload
	End Sub
	
	Public Sub setOrderNotes(strOrderNotes)
		orderNotes = strOrderNotes
	End Sub
	
	Public Sub setNoRegistration(strNoRegistration)
		noRegistration = strNoRegistration
	End Sub
	
	Public Sub setIdAdRef(strIdAdRef)
		idAdRef = strIdAdRef
	End Sub
		
 				
	Public Function insertOrdineNoTransaction(numIDUtente, dtaInserimento, strStatoOrd, numTotaleImponibile, numTotaleTasse, numTotale, strTipoPagam, dblPaymentCommission, bolPagamEffettuato, strGUID, userNotifiedXDownload, strOrderNotes, bolNoRegistration, numIdAds)
		on error resume next
		insertOrdineNoTransaction = -1
		
		Dim objDB, strSQL, strSQLSelect, objRS
		Dim objConn
		strSQL = "INSERT INTO ordini(id_utente, dta_inserimento, stato_ordine, totale_imponibile, totale_tasse, totale, tipo_pagam, payment_commission, pagam_effettuato, order_guid, user_notified_x_download, notes, no_registration, id_ads) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,?,?,?,?,?"
		if (not(isNull(numIdAds)) AND Trim(numIdAds)<>"") then
			strSQL = strSQL & ",?"
		else
			strSQL = strSQL & ",NULL"
		end if		
		strSQL = strSQL & ");"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDUtente)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtaInserimento)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStatoOrd)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotaleImponibile))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotaleTasse))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotale))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTipoPagam)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(dblPaymentCommission))
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolPagamEffettuato)		
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strGUID)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,userNotifiedXDownload)		
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strOrderNotes)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolNoRegistration)
		if (not(isNull(numIdAds)) AND Trim(numIdAds)<>"") then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIdAds)
		objCommand.Execute()
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(ordini.id_ordine) as id FROM ordini")
		if not (objRS.EOF) then
			insertOrdineNoTransaction = objRS("id")	
		end if			
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyOrdineNoTransaction(id, numIDUtente, dtaInserimento, strStatoOrd, numTotaleImponibile, numTotaleTasse, numTotale, strTipoPagam, dblPaymentCommission, bolPagamEffettuato, userNotifiedXDownload, strOrderNotes, bolNoRegistration, numIdAds)
		on error resume next
		Dim objDB, strSQL, objRS
		Dim objConn
		strSQL = "UPDATE ordini SET "
		strSQL = strSQL & "id_utente=?,"
		strSQL = strSQL & "dta_inserimento=?,"
		strSQL = strSQL & "stato_ordine=?,"
		strSQL = strSQL & "totale_imponibile=?," 
		strSQL = strSQL & "totale_tasse=?," 
		strSQL = strSQL & "totale=?," 
		strSQL = strSQL & "tipo_pagam=?," 
		strSQL = strSQL & "payment_commission=?," 
		strSQL = strSQL & "pagam_effettuato=?,"	
		strSQL = strSQL & "user_notified_x_download=?,"	
		strSQL = strSQL & "notes=?,"	
		strSQL = strSQL & "no_registration=?"
		if (not(isNull(numIdAds)) AND Trim(numIdAds)<>"") then
			strSQL = strSQL & ",id_ads=?"
		end if
		strSQL = strSQL & " WHERE id_ordine=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDUtente)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtaInserimento)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStatoOrd)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotaleImponibile))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotaleTasse))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotale))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTipoPagam)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(dblPaymentCommission))
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolPagamEffettuato)	
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,userNotifiedXDownload)		
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strOrderNotes)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolNoRegistration)
		if (not(isNull(numIdAds)) AND Trim(numIdAds)<>"") then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIdAds)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub				
 				
	Public Function insertOrdine(numIDUtente, dtaInserimento, strStatoOrd, numTotaleImponibile, numTotaleTasse, numTotale, strTipoPagam, dblPaymentCommission, bolPagamEffettuato, strGUID, userNotifiedXDownload, strOrderNotes, bolNoRegistration, numIdAds, objConn)
		on error resume next
		insertOrdine = -1
		
		Dim objDB, strSQL, strSQLSelect, objRS
		strSQL = "INSERT INTO ordini(id_utente, dta_inserimento, stato_ordine, totale_imponibile, totale_tasse, totale, tipo_pagam, payment_commission, pagam_effettuato, order_guid, user_notified_x_download, notes, no_registration, id_ads) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,?,?,?,?,?"
		if (not(isNull(numIdAds)) AND Trim(numIdAds)<>"") then
			strSQL = strSQL & ",?"
		else
			strSQL = strSQL & ",NULL"
		end if		
		strSQL = strSQL & ");"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDUtente)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtaInserimento)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStatoOrd)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotaleImponibile))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotaleTasse))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotale))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTipoPagam)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(dblPaymentCommission))
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolPagamEffettuato)		
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strGUID)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,userNotifiedXDownload)		
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strOrderNotes)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolNoRegistration)
		if (not(isNull(numIdAds)) AND Trim(numIdAds)<>"") then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIdAds)
		objCommand.Execute()
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(ordini.id_ordine) as id FROM ordini")
		if not (objRS.EOF) then
			insertOrdine = objRS("id")	
		end if			
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyOrdine(id, numIDUtente, dtaInserimento, strStatoOrd, numTotaleImponibile, numTotaleTasse, numTotale, strTipoPagam, dblPaymentCommission, bolPagamEffettuato, userNotifiedXDownload, strOrderNotes, bolNoRegistration, numIdAds, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE ordini SET "
		strSQL = strSQL & "id_utente=?,"
		strSQL = strSQL & "dta_inserimento=?,"
		strSQL = strSQL & "stato_ordine=?,"
		strSQL = strSQL & "totale_imponibile=?," 
		strSQL = strSQL & "totale_tasse=?," 
		strSQL = strSQL & "totale=?," 
		strSQL = strSQL & "tipo_pagam=?," 
		strSQL = strSQL & "payment_commission=?," 
		strSQL = strSQL & "pagam_effettuato=?,"	
		strSQL = strSQL & "user_notified_x_download=?,"	
		strSQL = strSQL & "notes=?,"	
		strSQL = strSQL & "no_registration=?"
		if (not(isNull(numIdAds)) AND Trim(numIdAds)<>"") then
			strSQL = strSQL & ",id_ads=?"
		end if
		strSQL = strSQL & " WHERE id_ordine=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numIDUtente)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtaInserimento)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStatoOrd)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotaleImponibile))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotaleTasse))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numTotale))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTipoPagam)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(dblPaymentCommission))
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolPagamEffettuato)	
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,userNotifiedXDownload)		
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strOrderNotes)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolNoRegistration)
		if (not(isNull(numIdAds)) AND Trim(numIdAds)<>"") then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIdAds)
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
	
	Public Sub changeStateOrder(id, strStatoOrd, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE ordini SET "
		strSQL = strSQL & "stato_ordine=?"				
		strSQL = strSQL & " WHERE id_ordine=?;" 

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStatoOrd)
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
	
	Public Sub changeStateOrderNoTransaction(id, strStatoOrd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE ordini SET "
		strSQL = strSQL & "stato_ordine=?"				
		strSQL = strSQL & " WHERE id_ordine=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strStatoOrd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
	
	Public Sub changePagamDoneOrder(id, numPagamDone, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE ordini SET "
		strSQL = strSQL & "pagam_effettuato=?"				
		strSQL = strSQL & " WHERE id_ordine=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numPagamDone)	
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
	
	Public Sub changePagamDoneOrderNoTransaction(id, numPagamDone)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE ordini SET "
		strSQL = strSQL & "pagam_effettuato=?"				
		strSQL = strSQL & " WHERE id_ordine=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numPagamDone)	
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
	
	Public Sub changeUserNotifiedOrder(id, bolNotified, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE ordini SET "
		strSQL = strSQL & "user_notified_x_download=?"				
		strSQL = strSQL & " WHERE id_ordine=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolNotified)	
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
	
	Public Sub changeUserNotifiedOrderNoTransaction(id, bolNotified)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE ordini SET "
		strSQL = strSQL & "user_notified_x_download=?"				
		strSQL = strSQL & " WHERE id_ordine=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolNotified)	
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteOrdineNoTransaction(id)
		on error resume next
		Dim objDB, strSQLDelField4ProdOrdine, strSQLDelProdOrdine, strSQLDelOrdine, strSQLDelSpeseOrdine, strSQLDelShipOrdine, objRS, objConn		 
		strSQLDelField4ProdOrdine = "DELETE FROM product_fields_x_order WHERE id_order=?;" 
		strSQLDelSpeseOrdine = "DELETE FROM spese_x_ordine WHERE id_ordine=?;" 
		strSQLDelProdOrdine = "DELETE FROM prodotti_x_ordine WHERE id_ordine=?;" 
		strSQLDelShipOrdine = "DELETE FROM order_shipping_address WHERE id_order=?;"
		strSQLDelRuleOrdine = "DELETE FROM business_rules_x_ordine WHERE id_order=?;" 
		strSQLDelVoucherOrdine = "DELETE FROM voucher_x_ordine WHERE id_order=?;" 
		strSQLDelOrdine = "DELETE FROM ordini WHERE id_ordine=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		Dim objCommand, objCommand2, objCommand3, objCommand4, objCommand5, objCommand6, objCommand7
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		Set objCommand3 = Server.CreateObject("ADODB.Command")
		Set objCommand4 = Server.CreateObject("ADODB.Command")
		Set objCommand5 = Server.CreateObject("ADODB.Command")
		Set objCommand6 = Server.CreateObject("ADODB.Command")
		Set objCommand7 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand3.ActiveConnection = objConn
		objCommand4.ActiveConnection = objConn
		objCommand5.ActiveConnection = objConn
		objCommand6.ActiveConnection = objConn
		objCommand7.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand3.CommandType=1
		objCommand4.CommandType=1
		objCommand5.CommandType=1
		objCommand6.CommandType=1
		objCommand7.CommandType=1
		objCommand.CommandText = strSQLDelField4ProdOrdine
		objCommand2.CommandText = strSQLDelSpeseOrdine
		objCommand3.CommandText = strSQLDelProdOrdine
		objCommand4.CommandText = strSQLDelShipOrdine
		objCommand5.CommandText = strSQLDelRuleOrdine
		objCommand6.CommandText = strSQLDelVoucherOrdine
		objCommand7.CommandText = strSQLDelOrdine
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)
		objCommand4.Parameters.Append objCommand4.CreateParameter(,19,1,,id)
		objCommand5.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand6.Parameters.Append objCommand6.CreateParameter(,19,1,,id)
		objCommand7.Parameters.Append objCommand6.CreateParameter(,19,1,,id)

		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand.Execute()
			objCommand2.Execute()
			objCommand3.Execute()
			objCommand4.Execute()
			objCommand5.Execute()
			objCommand6.Execute()
		end if
		objCommand7.Execute()

		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		Set objCommand4 = Nothing
		Set objCommand5 = Nothing
		Set objCommand6 = Nothing
		Set objCommand7 = Nothing
		
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
		
	Public Sub deleteOrdine(id, objConn)
		on error resume next
		Dim objDB, strSQLDelField4ProdOrdine, strSQLDelProdOrdine, strSQLDelOrdine, objRS, strSQLDelSpeseOrdine, strSQLDelShipOrdine
		strSQLDelField4ProdOrdine = "DELETE FROM product_fields_x_order WHERE id_order=?;" 
		strSQLDelSpeseOrdine = "DELETE FROM spese_x_ordine WHERE id_ordine=?;" 
		strSQLDelProdOrdine = "DELETE FROM prodotti_x_ordine WHERE id_ordine=?;" 
		strSQLDelShipOrdine = "DELETE FROM order_shipping_address WHERE id_order=?;" 
		strSQLDelRuleOrdine = "DELETE FROM business_rules_x_ordine WHERE id_order=?;" 
		strSQLDelVoucherOrdine = "DELETE FROM voucher_x_ordine WHERE id_order=?;" 
		strSQLDelOrdine = "DELETE FROM ordini WHERE id_ordine=?;"

		Dim objCommand, objCommand2, objCommand3, objCommand4, objCommand5, objCommand6, objCommand7
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		Set objCommand3 = Server.CreateObject("ADODB.Command")
		Set objCommand4 = Server.CreateObject("ADODB.Command")
		Set objCommand5 = Server.CreateObject("ADODB.Command")
		Set objCommand6 = Server.CreateObject("ADODB.Command")
		Set objCommand7 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand3.ActiveConnection = objConn
		objCommand4.ActiveConnection = objConn
		objCommand5.ActiveConnection = objConn
		objCommand6.ActiveConnection = objConn
		objCommand7.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand3.CommandType=1
		objCommand4.CommandType=1
		objCommand5.CommandType=1
		objCommand6.CommandType=1
		objCommand7.CommandType=1
		objCommand.CommandText = strSQLDelField4ProdOrdine
		objCommand2.CommandText = strSQLDelSpeseOrdine
		objCommand3.CommandText = strSQLDelProdOrdine
		objCommand4.CommandText = strSQLDelShipOrdine
		objCommand5.CommandText = strSQLDelRuleOrdine
		objCommand6.CommandText = strSQLDelVoucherOrdine
		objCommand7.CommandText = strSQLDelOrdine
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)
		objCommand4.Parameters.Append objCommand4.CreateParameter(,19,1,,id)
		objCommand5.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand6.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand7.Parameters.Append objCommand5.CreateParameter(,19,1,,id)

		if(Application("use_innodb_table") = 0) then
			objCommand.Execute()
			objCommand2.Execute()
			objCommand3.Execute()
			objCommand4.Execute()
			objCommand5.Execute()
			objCommand6.Execute()
		end if
		objCommand7.Execute()

		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		Set objCommand4 = Nothing
		Set objCommand5 = Nothing
		Set objCommand6 = Nothing
		Set objCommand7 = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Function findOrdini(id_order, id_user, dta_ins_from, dta_ins_to, stato_order, tipo_pagam, pagam_done, order_by, details, order_guid)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objOrdine
		findOrdini = null  
		strSQL = "SELECT * FROM ordini"

		if (isNull(id_order) AND isNull(id_user) AND isNull(dta_ins_from)  AND isNull(dta_ins_to) AND isNull(stato_order) AND isNull(tipo_pagam) AND isNull(pagam_done) AND isNull(order_guid)) then
			strSQL = "SELECT * FROM ordini"
		else
			strSQL = strSQL & " WHERE"
			
			if not(isNull(id_order)) then strSQL = strSQL & " AND id_ordine=?"
			if not(isNull(id_user)) then strSQL = strSQL & " AND id_utente =?"
			if not(isNull(dta_ins_from)) then
				DD = DatePart("d", dta_ins_from)
				MM = DatePart("m", dta_ins_from)
				YY = DatePart("yyyy", dta_ins_from)
				HH = 00
				MIN = 00
				SS = 00
				dta_ins_from = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS
				
				strSQL = strSQL & " AND dta_inserimento >=?" 
			end if
			if not(isNull(dta_ins_to)) then
				DD = DatePart("d", dta_ins_to)
				MM = DatePart("m", dta_ins_to)
				YY = DatePart("yyyy", dta_ins_to)
				HH = 23
				MIN = 59
				SS = 59
				dta_ins_to = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS
				
				strSQL = strSQL & " AND dta_inserimento <=?" 
			end if
			if not(isNull(stato_order)) then strSQL = strSQL & " AND stato_ordine =?"
			if not(isNull(tipo_pagam)) then strSQL = strSQL & " AND tipo_pagam =?"
			if not(isNull(pagam_done)) then strSQL = strSQL & " AND pagam_effettuato=?"
			if not(isNull(order_guid)) then strSQL = strSQL & " AND order_guid=?"
		end if
		
		Select Case order_by
		   Case 1
		      strSQL = strSQL & " ORDER BY id_ordine ASC"
		   Case 2
		      strSQL = strSQL & " ORDER BY id_ordine DESC"
		   Case 3
		      strSQL = strSQL & " ORDER BY id_utente ASC"
		   Case 4
		      strSQL = strSQL & " ORDER BY id_utente DESC"
		   Case 5
		      strSQL = strSQL & " ORDER BY dta_inserimento ASC"
		   Case 6
		      strSQL = strSQL & " ORDER BY dta_inserimento DESC"
		   Case 7
		      strSQL = strSQL & " ORDER BY stato_ordine ASC"
		   Case 8
		      strSQL = strSQL & " ORDER BY stato_ordine DESC"
		   Case 9
		      strSQL = strSQL & " ORDER BY totale ASC"
		   Case 10
		      strSQL = strSQL & " ORDER BY totale DESC"
		   Case 11
		      strSQL = strSQL & " ORDER BY tipo_pagam ASC"
		   Case 12
		      strSQL = strSQL & " ORDER BY tipo_pagam DESC"
		   Case 13
		      strSQL = strSQL & " ORDER BY pagam_effettuato ASC"
		   Case 14
		      strSQL = strSQL & " ORDER BY pagam_effettuato DESC"
		   Case Else
		      strSQL = strSQL & " ORDER BY id_ordine ASC"
		End Select

		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		if (isNull(id_order) AND isNull(id_user) AND isNull(dta_ins_from) AND isNull(dta_ins_to) AND isNull(stato_order) AND isNull(tipo_pagam) AND isNull(pagam_done) AND isNull(order_guid)) then
		else
			if not(isNull(id_order)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
			if not(isNull(id_user)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user)
			if not(isNull(dta_ins_from)) then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins_from)
			if not(isNull(dta_ins_to)) then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins_to)
			if not(isNull(stato_order)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,stato_order)
			if not(isNull(tipo_pagam)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,tipo_pagam)
			if not(isNull(pagam_done)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,pagam_done)	
			if not(isNull(order_guid)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,order_guid)
		end if
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then
			Dim objProdotti, objListaProdotti
			Set objProdotti = new Products4OrderClass   
			Set objDict = Server.CreateObject("Scripting.Dictionary")   
			do while not objRS.EOF
				Set objOrdine = new OrderClass
				strID = objRS("id_ordine")				
				
				If(CBool(details)) then
					Set objListaProdotti = objProdotti.getListaProdottiXOrdine(strID)				
					if not(isEmpty(objListaProdotti)) then
						objOrdine.setProdottiXOrdine(objListaProdotti)
						Set objListaProdotti = nothing
					else
						Set objListaProdotti = nothing
					end if				
				end if
				
				objOrdine.setIDOrdine(objRS("id_ordine"))
				objOrdine.setIDUtente(objRS("id_utente"))    
				objOrdine.setDtaInserimento(objRS("dta_inserimento"))
				objOrdine.setStatoOrdine(objRS("stato_ordine"))
				objOrdine.setTotaleImponibile(objRS("totale_imponibile"))
				objOrdine.setTotaleTasse(objRS("totale_tasse"))
				objOrdine.setTotale(objRS("totale"))
				objOrdine.setTipoPagam(objRS("tipo_pagam"))
				objOrdine.setPaymentCommission(objRS("payment_commission"))
				objOrdine.setPagamEffettuato(objRS("pagam_effettuato"))
				objOrdine.setOrderGUID(objRS("order_guid"))
				objOrdine.setUserNotifiedXDownload(objRS("user_notified_x_download"))
				objOrdine.setOrderNotes(objRS("notes"))
				objOrdine.setNoRegistration(objRS("no_registration"))
				objOrdine.setIdAdRef(objRS("id_ads"))				
				
				objDict.add strID, objOrdine
				Set objOrdine = nothing
				objRS.moveNext()
			loop
			
			Set objProdotti = Nothing
			Set findOrdini = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
		
	Public Function getListaOrdini(order_by, details)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objOrdine
		getListaOrdini = null  
		strSQL = "SELECT * FROM ordini"
		
		Select Case order_by
		   Case 1
		      strSQL = strSQL & " ORDER BY dta_inserimento ASC"
		   Case 2
		      strSQL = strSQL & " ORDER BY dta_inserimento DESC"
		   Case 3
		      strSQL = strSQL & " ORDER BY totale ASC"
		   Case 4
		      strSQL = strSQL & " ORDER BY totale DESC"
		   Case Else
		      strSQL = strSQL & " ORDER BY id_ordine ASC"
		End Select
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Set objRS = objConn.Execute(strSQL)
		
		if not(objRS.EOF) then
			Dim objProdotti, objListaProdotti
			Set objProdotti = new Products4OrderClass   
			Set objDict = Server.CreateObject("Scripting.Dictionary")   
			do while not objRS.EOF
				Set objOrdine = new OrderClass
				strID = objRS("id_ordine")				
				
				If(CBool(details)) then
					Set objListaProdotti = objProdotti.getListaProdottiXOrdine(strID)				
					if not(isEmpty(objListaProdotti)) then
						objOrdine.setProdottiXOrdine(objListaProdotti)
						Set objListaProdotti = nothing
					else
						Set objListaProdotti = nothing
					end if				
				end if
				
				objOrdine.setIDOrdine(objRS("id_ordine"))
				objOrdine.setIDUtente(objRS("id_utente"))    
				objOrdine.setDtaInserimento(objRS("dta_inserimento"))
				objOrdine.setStatoOrdine(objRS("stato_ordine"))
				objOrdine.setTotaleImponibile(objRS("totale_imponibile"))
				objOrdine.setTotaleTasse(objRS("totale_tasse"))
				objOrdine.setTotale(objRS("totale"))
				objOrdine.setTipoPagam(objRS("tipo_pagam"))
				objOrdine.setPaymentCommission(objRS("payment_commission"))
				objOrdine.setPagamEffettuato(objRS("pagam_effettuato"))
				objOrdine.setOrderGUID(objRS("order_guid"))
				objOrdine.setUserNotifiedXDownload(objRS("user_notified_x_download"))
				objOrdine.setOrderNotes(objRS("notes"))
				objOrdine.setNoRegistration(objRS("no_registration"))
				objOrdine.setIdAdRef(objRS("id_ads"))
				
				objDict.add strID, objOrdine
				Set objOrdine = nothing
				objRS.moveNext()
			loop
			
			Set objProdotti = Nothing
			Set getListaOrdini = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
			
	Public Function findOrdineByID(id, details)
		'on error resume next
		
		findOrdineByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM ordini WHERE id_ordine=?;"
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		Set objRS = objCommand.Execute()		
		
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")		
		else
			Dim objOrdine, tmpID
			Dim objProdotti, objListaProdotti
			Set objProdotti = new Products4OrderClass 
			Set objOrdine = new OrderClass	
			tmpID = objRS("id_ordine")
			
			If(CBool(details)) then	
				On Error Resume Next
				Set objListaProdotti = objProdotti.getListaProdottiXOrdine(tmpID)		
				if Err.number = 0 then	
					if not(isEmpty(objListaProdotti)) AND not isNull(objListaProdotti) then
						objOrdine.setProdottiXOrdine(objListaProdotti)
						Set objListaProdotti = nothing
					else
						Set objListaProdotti = nothing
					end if				
				end if
			end if

			objOrdine.setIDOrdine(tmpID)
			objOrdine.setIDUtente(objRS("id_utente"))    
			objOrdine.setDtaInserimento(objRS("dta_inserimento"))
			objOrdine.setStatoOrdine(objRS("stato_ordine"))
			objOrdine.setTotaleImponibile(objRS("totale_imponibile"))
			objOrdine.setTotaleTasse(objRS("totale_tasse"))
			objOrdine.setTotale(objRS("totale"))
			objOrdine.setTipoPagam(objRS("tipo_pagam"))
			objOrdine.setPaymentCommission(objRS("payment_commission"))
			objOrdine.setPagamEffettuato(objRS("pagam_effettuato"))
			objOrdine.setOrderGUID(objRS("order_guid"))
			objOrdine.setUserNotifiedXDownload(objRS("user_notified_x_download"))
			objOrdine.setOrderNotes(objRS("notes"))
			objOrdine.setNoRegistration(objRS("no_registration"))
			objOrdine.setIdAdRef(objRS("id_ads"))
									
			Set objProdotti = Nothing
			Set findOrdineByID = objOrdine
			Set objOrdine = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Function getMaxIDOrdine()
		on error resume next
		
		getMaxIDOrdine = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT MAX(id_ordine) AS id_ord FROM ordini;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDOrdine = objRS("id_ord")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Function verifyOrder(id, guid, amount)
		on error resume next
		
		verifyOrder = false
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT id_ordine FROM ordini WHERE id_ordine=? AND totale=? AND order_guid=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(amount))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,guid)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			verifyOrder = false		
		else
			verifyOrder = true	
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Function countUserOrder(id_user)
		on error resume next
		
		countUserOrder = 0
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT count(id_ordine) as counter FROM ordini WHERE id_utente=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			countUserOrder = 0		
		else
			countUserOrder = objRS("counter")	
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function getListaStatiOrder()
		Set getListaStatiOrder = Server.CreateObject("Scripting.Dictionary")
		getListaStatiOrder.add "1", "backend.ordini.lista.table.select.option.ord_inserting"
		getListaStatiOrder.add "2", "backend.ordini.lista.table.select.option.ord_executing"
		getListaStatiOrder.add "3", "backend.ordini.lista.table.select.option.ord_executed"
		getListaStatiOrder.add "4", "backend.ordini.lista.table.select.option.ord_sca"
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