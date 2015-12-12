<%
Class AdsClass
	Private id_ads
	Private id_element
	Private cod_element
	Private id_utente
	Private phone
	Private ads_type
	Private price
	Private active
	Private dta_ins
	
	
	Public Function getID()
		getID = id_ads
	End Function
				
	Public Sub setID(numID)
		id_ads = numID
	End Sub
		
	Public Function getIDElement()
		getIDElement = id_element
	End Function
					
	Public Sub setIDElement(numIDElement)
		id_element = numIDElement
	End Sub
		
	Public Function getCodElement()
		getCodElement = cod_element
	End Function
					
	Public Sub setCodElement(numCodElement)
		cod_element = numCodElement
	End Sub
		
	Public Function getIDUtente()
		getIDUtente = id_utente
	End Function
					
	Public Sub setIDUtente(numIDUtente)
		id_utente = numIDUtente
	End Sub
		
	Public Function getPhone()
		getPhone = phone
	End Function
		
	Public Sub setPhone(strPhone)
		phone = strPhone
	End Sub
		
	Public Function getAdsType()
		getAdsType = ads_type
	End Function
		
	Public Sub setAdsType(strAdsType)
		ads_type = strAdsType
	End Sub
		
	Public Function getPrice()
		getPrice = price
	End Function
		
	Public Sub setPrice(strPrice)
		price = strPrice
	End Sub
		
	Public Function isActive()
		isActive = active
	End Function
		
	Public Sub setActive(bolActive)
		active = bolActive
	End Sub
		
	Public Function getDtaInserimento()
		getDtaInserimento = dta_ins
	End Function
		
	Public Sub setDtaInserimento(dtaInserimento)
		dta_ins = dtaInserimento
	End Sub
	


'*********************************** METODI ADS *********************** 				
	Public Function insertAds(id_element, id_utente, strPhone, ads_type, price, dta_ins, objConn)
		on error resume next
		Dim strSQL, strSQLSelect, objRS
		
		insertAds = -1
				
		strSQL = "INSERT INTO ads(id_element, id_utente, phone, ads_type, price, dta_inserimento) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strPhone)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,ads_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(price))
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(ads.id_ads) as id FROM ads")
		if not (objRS.EOF) then
			insertAds = objRS("id")	
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
	
	Public Function insertAdsNoTransaction(id_element, id_utente, strPhone, ads_type, price, dta_ins)
		on error resume next
		Dim objDB, strSQL, strSQLSelect, objRS, objConn
		
		insertAdsNoTransaction = -1
				
		strSQL = "INSERT INTO ads(id_element, id_utente, phone, ads_type, price, dta_inserimento) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strPhone)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,ads_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(price))
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(ads.id_ads) as id FROM ads")
		if not (objRS.EOF) then
			insertAdsNoTransaction = objRS("id")	
		end if			
		Set objRS = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyAds(id_ads, strPhone, ads_type, price, objConn)
		on error resume next
		Dim strSQL, objRS
		strSQL = "UPDATE ads SET "	 
		strSQL = strSQL & "phone=?,"
		strSQL = strSQL & "ads_type=?,"
		strSQL = strSQL & "price=?"
		strSQL = strSQL & " WHERE id_ads=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strPhone)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,ads_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(price))
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ads)
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
		
	Public Sub modifyAdsNoTransaction(id_ads, strPhone, ads_type, price)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE ads SET "	 
		strSQL = strSQL & "phone=?,"
		strSQL = strSQL & "ads_type=?,"
		strSQL = strSQL & "price=?"
		strSQL = strSQL & " WHERE id_ads=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strPhone)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,ads_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(price))
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ads)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteAds(id)
		on error resume next
		Dim objDB, strSQLDelAds, strSQLDelAdsPromotion, objRS, objConn
		strSQLDelAds = "DELETE FROM ads WHERE id_ads=?;" 
		strSQLDelAdsPromotion = "DELETE FROM ads_promotion WHERE id_ads=?;"		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand, objCommand2
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQLDelAds
		objCommand2.CommandText = strSQLDelAdsPromotion
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		objCommand2.Parameters.Append objCommand.CreateParameter(,20,1,,id)		
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
				
	Public Function findAds(id_utente, id_element, ads_type, price_from, price_to, dta_ins_from, dta_ins_to, title, arrTargetCat, arrTargetLang)
		on error resume next		
		findAds = null
		
		Dim noTargetCat,noTargetlang, hasTarget
		hasTarget = true
		noTargetCat = (isNull(arrTargetCat) OR not(strComp(typename(arrTargetCat), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0))		

		if (noTargetCat) AND (noTargetlang) then
			hasTarget = false
		elseif (noTargetCat) OR (noTargetlang) then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=026")
		end if
		
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "SELECT ads.*, news.titolo FROM ads LEFT JOIN news ON ads.id_element=news.id WHERE stato_news=1"
					
		if not(isNull(ads_type)) AND not(Trim(ads_type)="") then strSQL = strSQL & " AND ads_type=?"
		if not(isNull(id_utente)) AND not(Trim(id_utente)="") then strSQL = strSQL & " AND id_utente=?"
		if not(isNull(id_element)) AND not(Trim(id_element)="") then strSQL = strSQL & " AND id_element=?"
		if not(isNull(price_from)) AND not(Trim(price_from)="") then strSQL = strSQL & " AND price >=?"
		if not(isNull(price_to)) AND not(Trim(price_to)="") then strSQL = strSQL & " AND price <=?"
		if not(isNull(dta_ins_from)) AND not(Trim(dta_ins_from)="") then strSQL = strSQL & " AND dta_inserimento >=?"
		if not(isNull(dta_ins_to)) AND not(Trim(dta_ins_to)="") then strSQL = strSQL & " AND dta_inserimento <=?"
		if not(isNull(title)) AND not(Trim(title)="") then strSQL = strSQL & " AND titolo LIKE ?"
		if (hasTarget) then 
			strSQL = strSQL & " AND news.id IN("					
			strSQL = strSQL & "SELECT DISTINCT(id_news) FROM target_x_news WHERE id_news IN("
			strSQL = strSQL & "SELECT DISTINCT(id_news) FROM target_x_news WHERE id_target IN("								
			for each idx in arrTargetCat
				strSQL = strSQL &idx&","
			next					
			strSQL = strSQL & "))"				
			strSQL = strSQL & " AND id_target IN("	
			for each idy in arrTargetLang
				strSQL = strSQL &idy&","
			next					
			strSQL = strSQL & "))"						
			strSQL = Replace(strSQL, ",)", ")", 1, -1, 1)
			strSQL = Trim(strSQL)
		end if
		
		strSQL = strSQL & " ORDER BY dta_inserimento DESC"
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"		
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if (isNull(ads_type) AND isNull(id_utente) AND isNull(dta_ins)) then
		else		
			if not(isNull(ads_type)) AND not(Trim(ads_type)="") then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,ads_type)
			if not(isNull(id_utente)) AND not(Trim(id_utente)="") then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
			if not(isNull(id_element)) AND not(Trim(id_element)="") then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
			if not(isNull(price_from)) AND not(Trim(price_from)="") then objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(price_from))
			if not(isNull(price_to)) AND not(Trim(price_to)="") then objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(price_to))
			if not(isNull(dta_ins_from)) AND not(Trim(dta_ins_from)="") then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins_from&" 00:00:00")
			if not(isNull(dta_ins_to)) AND not(Trim(dta_ins_to)="") then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins_to&" 23:59:59")
			if not(isNull(title)) AND not(Trim(title)="") then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,"%"&title&"%")
		end if
		
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then						
			Dim objAdsTmp, objDict, strIDTmp				
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
					
			do while not objRS.EOF
				Set objAdsTmp = New AdsClass
				strIDTmp = objRS("id_ads")
				objAdsTmp.setID(strIDTmp)
				objAdsTmp.setIDElement(objRS("id_element")) 
				objAdsTmp.setIDUtente(objRS("id_utente")) 
				objAdsTmp.setPhone(objRS("phone"))
				objAdsTmp.setAdsType(objRS("ads_type"))
				objAdsTmp.setPrice(objRS("price"))
				objAdsTmp.setDtaInserimento(objRS("dta_inserimento"))
				objDict.add strIDTmp, objAdsTmp
				Set objAdsTmp = Nothing
				objRS.moveNext()
			loop
						
			Set findAds = objDict
			Set objDict = nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findAdByID(id_ads)
		on error resume next		
		findAdByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM ads WHERE id_ads=?;" 
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ads)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then
			Set objAdsTmp = New AdsClass
			strIDTmp = objRS("id_ads")
			objAdsTmp.setID(strIDTmp)
			objAdsTmp.setIDElement(objRS("id_element")) 
			objAdsTmp.setIDUtente(objRS("id_utente")) 
			objAdsTmp.setPhone(objRS("phone"))
			objAdsTmp.setAdsType(objRS("ads_type"))
			objAdsTmp.setPrice(objRS("price"))
			objAdsTmp.setDtaInserimento(objRS("dta_inserimento"))						
			Set findAdByID = objAdsTmp
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findAdByElement(id_utente, id_element)
		on error resume next		
		findAdByElement = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM ads WHERE id_element=?"
		if not(isNull(id_utente)) then strSQL = strSQL & " AND id_utente=?"
		strSQL = strSQL & ";"		
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		if not(isNull(id_utente)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then
			Set objAdsTmp = New AdsClass
			strIDTmp = objRS("id_ads")
			objAdsTmp.setID(strIDTmp)
			objAdsTmp.setIDElement(objRS("id_element")) 
			objAdsTmp.setIDUtente(objRS("id_utente")) 
			objAdsTmp.setPhone(objRS("phone"))
			objAdsTmp.setAdsType(objRS("ads_type"))
			objAdsTmp.setPrice(objRS("price"))
			objAdsTmp.setDtaInserimento(objRS("dta_inserimento"))						
			Set findAdByElement = objAdsTmp
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function countAdsByIDUtente(id_utente)
		on error resume next		
		countAdsByIDUtente = 0
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT count(*) AS counter FROM ads WHERE id_utente=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_utente)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then							
			countAdsByIDUtente = (countAdsByIDUtente + Cint(objRS("counter")))
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function



	'*********************************************************** CODICE PER LA GESTIONE DEGLI ANNUNCI PROMOZIONALI A PAGAMENTO
	Public Sub insertAdsPromotion(id_ad, id_element, cod_element, active, dta_ins, objConn)
		on error resume next
		Dim objDB, strSQL, strSQLSelect, objRS
				
		strSQL = "INSERT INTO ads_promotion(id_ads, id_element, cod_element, active, dta_inserimento) VALUES("
		strSQL = strSQL & "?,?,?,?,?);"
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,cod_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
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
	
	Public Sub insertAdsPromotionNoTransaction(id_ad, id_element, cod_element, active, dta_ins)
		on error resume next
		Dim objDB, strSQL, strSQLSelect, objRS, objConn
				
		strSQL = "INSERT INTO ads_promotion(id_ads, id_element, cod_element, active, dta_inserimento) VALUES("
		strSQL = strSQL & "?,?,?,?,?);"
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,cod_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
		objCommand.Execute()
		Set objCommand = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyAdsPromotion(id_ad, id_element, cod_element, active, dta_ins, objConn)
		on error resume next
		Dim strSQL, objRS
		strSQL = "UPDATE ads_promotion SET "	 
		strSQL = strSQL & "cod_element=?,"
		strSQL = strSQL & "active=?"
		if not(isNull(dta_ins)) then strSQL = strSQL & ",dta_inserimento=?"
		strSQL = strSQL & " WHERE id_ads=? AND id_element=?;"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,cod_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active)
		if not(isNull(dta_ins)) then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
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
		
	Public Sub modifyAdsPromotionNoTransaction(id_ad, id_element, cod_element, active, dta_ins)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE ads_promotion SET "	 
		strSQL = strSQL & "cod_element=?,"
		strSQL = strSQL & "active=?"
		if not(isNull(dta_ins)) then strSQL = strSQL & ",dta_inserimento=?"
		strSQL = strSQL & " WHERE id_ads=? AND id_element=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,cod_element)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active)
		if not(isNull(dta_ins)) then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub activateAdsPromotion(id_ad, id_element, objConn)
		on error resume next
		Dim strSQL, objRS
		strSQL = "UPDATE ads_promotion SET "	
		strSQL = strSQL & "active=1"
		strSQL = strSQL & ",dta_inserimento=?"
		strSQL = strSQL & " WHERE id_ads=? AND id_element=?;"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,now())
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_element)
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
		
	Public Sub activateAdsPromotionNoTransaction(id_ad, id_element)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE ads_promotion SET "	
		strSQL = strSQL & "active=1"
		strSQL = strSQL & ",dta_inserimento=?"
		strSQL = strSQL & " WHERE id_ads=? AND id_element=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,now())
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_element)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteAdsPromotion(id_ad,id_element, objConn)
		on error resume next
		Dim strSQLDelAds
		strSQLDelAds = "DELETE FROM ads_promotion WHERE id_ads=? AND id_element=?;" 	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDelAds
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_element)
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
		
	Public Sub deleteAdsPromotionNoTransaction(id_ad,id_element)
		on error resume next
		Dim objDB, strSQLDelAds, objConn
		strSQLDelAds = "DELETE FROM ads_promotion WHERE id_ads=? AND id_element=?;" 		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDelAds
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_element)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
				
	Public Function findAdsPromotionByID(id_ad)
		on error resume next		
		findAdsPromotionByID = null
		
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "SELECT * FROM ads_promotion WHERE id_ads=?;"	
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ad)
		
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then						
			Dim objAdsTmp, objDict, strIDTmp				
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
					
			do while not objRS.EOF
				Set objAdsTmp = New AdsClass
				strIDTmp = objRS("id_ads")
				strIDELTmp = objRS("id_element")
				objAdsTmp.setID(strIDTmp)
				objAdsTmp.setIDElement(strIDELTmp) 
				objAdsTmp.setCodElement(objRS("cod_element")) 
				objAdsTmp.setActive(objRS("active"))
				objAdsTmp.setDtaInserimento(objRS("dta_inserimento"))
				objDict.add strIDTmp&"#"&strIDELTmp, objAdsTmp
				Set objAdsTmp = Nothing
				objRS.moveNext()
			loop
						
			Set findAdsPromotionByID = objDict
			Set objDict = nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findAdPromotionByElement(id_ad, id_element)
		on error resume next		
		findAdPromotionByElement = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM ads_promotion WHERE id_ads=? AND id_element=?;" 
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_ad)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_element)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then
			Set objAdsTmp = New AdsClass
			strIDTmp = objRS("id_ads")
			objAdsTmp.setID(strIDTmp)
			objAdsTmp.setIDElement(objRS("id_element")) 
			objAdsTmp.setCodElement(objRS("cod_element")) 
			objAdsTmp.setActive(objRS("active"))
			objAdsTmp.setDtaInserimento(objRS("dta_inserimento"))					
			Set findAdPromotionByElement = objAdsTmp
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = Replace(doubleValue, ".",",")			
	End Function
	
	'public Sub toString()
		'response.write ()
	'end Sub
End Class
%>