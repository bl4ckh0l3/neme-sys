<%
Class NewsClass
	Private news_id
	Private titolo
	Private abstract1
	Private abstract2
	Private abstract3
	Private testo
	Private keyword
	Private data_inserimento
	Private data_pubblicazione
	Private data_cancellazione
	Private stato_news
	Private user_Editor_ID
	Private objFileNews
	Private objTargetNews
	Private objFieldsNews
	Private meta_description
	Private meta_keyword
	Private page_title
	
	
	Public Function getNewsID()
		getNewsID = news_id
	End Function
	
	Public Sub setNewsID(strID)
		news_id = strID
	End Sub	
	
	Public Function getTitolo()
		getTitolo = titolo
	End Function
	
	Public Sub setTitolo(strTitolo)
		titolo = strTitolo
	End Sub
	
	Public Function getAbstract1()
		getAbstract1 = abstract1
	End Function
	
	Public Sub setAbstract1(strAbs1)
		abstract1 = strAbs1
	End Sub
	
	Public Function getAbstract2()
		getAbstract2 = abstract2
	End Function
	
	Public Sub setAbstract2(strAbs2)
		abstract2 = strAbs2
	End Sub
	
	Public Function getAbstract3()
		getAbstract3 = abstract3
	End Function
	
	Public Sub setAbstract3(strAbs3)
		abstract3 = strAbs3
	End Sub
	
	Public Function getTesto()
		getTesto = testo
	End Function
	
	Public Sub setTesto(strTesto)
		testo = strTesto
	End Sub
	
	Public Function getKeyword()
		getKeyword = keyword
	End Function
	
	Public Sub setKeyword(strKeyword)
		keyword = strKeyword
	End Sub
	
	Public Function getDataInsNews()
		getDataInsNews = data_inserimento
	End Function
	
	Public Sub setDataInsNews(Data_ins)
		data_inserimento = Data_ins
	End Sub
	
	Public Function getDataPubNews()
		getDataPubNews = data_pubblicazione
	End Function
	
	Public Sub setDataPubNews(Data_pub)
		data_pubblicazione = Data_pub
	End Sub
	
	Public Function getDataDelNews()
		getDataDelNews = data_cancellazione
	End Function
	
	Public Sub setDataDelNews(Data_del)
		data_cancellazione = Data_del
	End Sub
		
	Public Function getStato()
		getStato = stato_news
	End Function
	
	Public Sub setStato(strStato)
		stato_news = strStato
	End Sub
	
	Public Function getEditorID()
		getEditorID = user_Editor_ID
	End Function
	
	Public Sub setEditorID(strEdID)
		user_Editor_ID = strEdID
	End Sub
	
	Public Function getListaTarget()
		Set getListaTarget = objTargetNews
	End Function
	
	Public Sub setListaTarget(objTarget)
		Set objTargetNews = objTarget
	End Sub
	
	Public Function getFilePerNews()	
		if(isNull(objFileNews) OR isEmpty(objFileNews)) then
			getFilePerNews = null
		else
			Set getFilePerNews = objFileNews
		end if
	End Function
	
	Public Sub setFilePerNews(objFiles)
		if(isNull(objFiles)) then
			objFileNews = null
		else
			Set objFileNews = objFiles
		end if		
	End Sub
	
	Public Function getListaFields()
		Set getListaFields = objFieldsNews
	End Function
	
	Public Sub setListaFields(objFieldsN)
		Set objFieldsNews = objFieldsN
	End Sub
	
	Public Function getMetaDescription()
		getMetaDescription = meta_description
	End Function
	
	Public Sub setMetaDescription(strMetaDesc)
		meta_description = strMetaDesc
	End Sub
	
	Public Function getMetaKeyword()
		getMetaKeyword = meta_keyword
	End Function
	
	Public Sub setMetaKeyword(strMetaKeyword)
		meta_keyword = strMetaKeyword
	End Sub
	
	Public Function getPageTitle()
		getPageTitle = page_title
	End Function
	
	Public Sub setPageTitle(strPageTitle)
		page_title = strPageTitle
	End Sub
	

	Public Function getMaxIDNews()
		on error resume next
		
		getMaxIDNews = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT max(news.id) as id FROM news;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDNews = objRS("id")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Function insertNews(strTitolo, strAbst1, strAbst2, strAbst3, strTesto, strKeyword, dtData_ins, dtData_pub, dtData_del, intStato, strMetaDesc, strMetaKey, strPageTitle, objConn)
		on error resume next
		Dim test
		insertNews = -1
		
		Dim objDB, strSQL, objRS, strDataDel
		
		if not(dtData_del = "") then
			strDataDel = convertDate(dtData_del)
		else
			'controllo il tipo di database in uso e imposto la data a null, se db Access
			if (Application("dbType") = 0) then
				strDataDel = "null"
			else		
				strDataDel = "'0000-00-00 00:00:00'"
			end if		
		end if

		if (Application("dbType") = 1) then
			dtData_ins = convertDate(dtData_ins)
			dtData_pub = convertDate(dtData_pub)
		end if			
		
		strSQL = "INSERT INTO news(titolo, abstract, abstract_2, abstract_3, testo, keyword, data_inserimento, data_pubblicazione, data_cancellazione, stato_news, meta_description, meta_keyword, page_title) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		if not(dtData_del = "") then
			strSQL = strSQL & "?,?,?,?,?);"
		else
			strSQL = strSQL & strDataDel&",?,?,?,?);"
		end if

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strTitolo)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst1)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst2)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst3)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strTesto)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKeyword)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_pub)
		if not(dtData_del = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,strDataDel)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,intStato)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Execute()
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(news.id) as id FROM news")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertNews = objRS("id")	
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
		
	Public Sub modifyNews(id, strTitolo, strAbst1, strAbst2, strAbst3, strTesto, strKeyword, dtData_ins, dtData_pub, dtData_del, intStato, strMetaDesc, strMetaKey, strPageTitle, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, strDataDel		

		if (Application("dbType") = 1) then
			dtData_ins = convertDate(dtData_ins)
			dtData_pub = convertDate(dtData_pub)
		end if

		if not(dtData_del = "") then
			strDataDel = convertDate(dtData_del)
		else	
			'controllo il tipo di database in uso e modifico il punto decimale con la virgola, se db Access
			if (Application("dbType") = 0) then
				strDataDel =  "null"
			else		
				strDataDel =  "'0000-00-00 00:00:00'"
			end if			
		end if

		strSQL = "UPDATE news SET "
		strSQL = strSQL & "titolo=?,"
		strSQL = strSQL & "abstract=?,"
		strSQL = strSQL & "abstract_2=?," 
		strSQL = strSQL & "abstract_3=?,"
		strSQL = strSQL & "testo=?,"
		strSQL = strSQL & "keyword=?,"
		strSQL = strSQL & "data_inserimento=?,"
		strSQL = strSQL & "data_pubblicazione=?,"
		if not(Trim(dtData_del) = "") then
			strSQL = strSQL & "data_cancellazione=?,"
		else
			strSQL = strSQL & "data_cancellazione="&strDataDel&","
		end if
		strSQL = strSQL & "stato_news=?,"
		strSQL = strSQL & "meta_description=?,"
		strSQL = strSQL & "meta_keyword=?,"
		strSQL = strSQL & "page_title=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strTitolo)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst1)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst2)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst3)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strTesto)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKeyword)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_pub)
		if not(Trim(dtData_del) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,strDataDel)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,intStato)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
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
			
	Public Function insertNewsNoTransaction(strTitolo, strAbst1, strAbst2, strAbst3, strTesto, strKeyword, dtData_ins, dtData_pub, dtData_del, intStato, strMetaDesc, strMetaKey, strPageTitle)
		on error resume next
		Dim test
		insertNewsNoTransaction = -1
		
		Dim objDB, strSQL, objRS, objConn, strDataDel
		
		if not(dtData_del = "") then
			strDataDel = convertDate(dtData_del)
		else
			'controllo il tipo di database in uso e imposto la data a null, se db Access
			if (Application("dbType") = 0) then
				strDataDel = "null"
			else		
				strDataDel = "'0000-00-00 00:00:00'"
			end if		
		end if

		if (Application("dbType") = 1) then
			dtData_ins = convertDate(dtData_ins)
			dtData_pub = convertDate(dtData_pub)
		end if			
		
		strSQL = "INSERT INTO news(titolo, abstract, abstract_2, abstract_3, testo, keyword, data_inserimento, data_pubblicazione, data_cancellazione, stato_news, meta_description, meta_keyword, page_title) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		if not(dtData_del = "") then
			strSQL = strSQL & "?,?,?,?,?);"
		else
			strSQL = strSQL & strDataDel&",?,?,?,?);"
		end if
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strTitolo)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst1)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst2)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst3)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strTesto)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKeyword)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_pub)
		if not(dtData_del = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,strDataDel)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,intStato)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Execute()
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(news.id) as id FROM news")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertNewsNoTransaction = objRS("id")	
		end if		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyNewsNoTransaction(id, strTitolo, strAbst1, strAbst2, strAbst3, strTesto, strKeyword, dtData_ins, dtData_pub, dtData_del, intStato, strMetaDesc, strMetaKey, strPageTitle)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, strDataDel	

		if (Application("dbType") = 1) then
			dtData_ins = convertDate(dtData_ins)
			dtData_pub = convertDate(dtData_pub)
		end if

		if not(dtData_del = "") then
			strDataDel = convertDate(dtData_del)
		else	
			'controllo il tipo di database in uso e modifico il punto decimale con la virgola, se db Access
			if (Application("dbType") = 0) then
				strDataDel =  "null"
			else		
				strDataDel =  "'0000-00-00 00:00:00'"
			end if			
		end if

		strSQL = "UPDATE news SET "
		strSQL = strSQL & "titolo=?,"
		strSQL = strSQL & "abstract=?,"
		strSQL = strSQL & "abstract_2=?," 
		strSQL = strSQL & "abstract_3=?,"
		strSQL = strSQL & "testo=?,"
		strSQL = strSQL & "keyword=?,"
		strSQL = strSQL & "data_inserimento=?,"
		strSQL = strSQL & "data_pubblicazione=?,"
		if not(dtData_del = "") then
			strSQL = strSQL & "data_cancellazione=?,"
		else
			strSQL = strSQL & "data_cancellazione="&strDataDel&","
		end if
		strSQL = strSQL & "stato_news=?,"
		strSQL = strSQL & "meta_description=?,"
		strSQL = strSQL & "meta_keyword=?,"
		strSQL = strSQL & "page_title=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strTitolo)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst1)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst2)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAbst3)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strTesto)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKeyword)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_ins)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtData_pub)
		if not(dtData_del = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,strDataDel)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,intStato)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
		
	Public Sub deleteNews(id)
		on error resume next
		Dim objDB, strSQLDelFile, strSQLDelUser, strSQLDelNews, strSQLDelTarget, strSQLDelLocaliz, strSQLDelAds, objRS, objConn
		
		strSQLDelFile = "DELETE FROM file_x_news WHERE id_news=?;"
		strSQLDelTarget = "DELETE FROM target_x_news WHERE id_news=?;"
		strSQLDelUser = "DELETE FROM news_x_utente WHERE id_news=?;"
		strSQLDelNews = "DELETE FROM news WHERE id=?;"
		strSQLDelLocaliz = "DELETE FROM googlemap_localization WHERE id_element=? AND `type`=1;"
		strSQLDelAds = "DELETE FROM ads WHERE id_element=?;"
		strSQLDelField = "DELETE FROM content_fields_match WHERE id_news=?;"
		
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
		objCommand.CommandText = strSQLDelFile
		objCommand2.CommandText = strSQLDelTarget
		objCommand3.CommandText = strSQLDelUser
		objCommand4.CommandText = strSQLDelNews
		objCommand5.CommandText = strSQLDelLocaliz
		objCommand6.CommandText = strSQLDelAds
		objCommand7.CommandText = strSQLDelField
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)
		objCommand4.Parameters.Append objCommand4.CreateParameter(,19,1,,id)
		objCommand5.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand6.Parameters.Append objCommand6.CreateParameter(,19,1,,id)
		objCommand7.Parameters.Append objCommand7.CreateParameter(,19,1,,id)

		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand.Execute()
			objCommand2.Execute()
			objCommand3.Execute()
			objCommand6.Execute()
			objCommand7.Execute()
		end if
		objCommand5.Execute()
		objCommand4.Execute()

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
	
	
	'****************** funzione di recupero delle news tramite vari parametri
	Public Function findNewsSlim(id, id_utente, titolo, keyword, arrTargetCat, arrTargetLang, data_pub, data_del, stato_news, order_by, bolAddTarget, bolAddFiles)
		on error resume next		
				
		findNewsSlim = null				
		Dim objDB, strSQL, strSQLTarget, strSQLTmp, objRS, objRSTargetCat, objRSTargetLang, objListTarget, objConn
		Dim hasTarget, doExit
		hasTarget = true

		Dim noTargetCat,noTargetlang
		noTargetCat = (isNull(arrTargetCat) OR not(strComp(typename(arrTargetCat), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0))
		
		Set objListTarget = Server.CreateObject("Scripting.Dictionary")
		if (noTargetCat) AND (noTargetlang) then
			hasTarget = false
		elseif (noTargetCat) OR (noTargetlang) then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=026")
		end if
		
		strSQL = "SELECT * FROM news_find"
		if (isNull(id) AND isNull(id_utente) AND isNull(titolo) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news) AND not(hasTarget)) then
			strSQL = "SELECT * FROM news_find"
		else				
			strSQL = strSQL & " WHERE"
			
			if not(isNull(id)) then strSQL = strSQL & " AND id=?"
			if not(isNull(id_utente)) then strSQL = strSQL & " AND id_utente=?"
			if not(isNull(titolo)) then strSQL = strSQL & " AND titolo=?"
			if not(isNull(keyword)) then strSQL = strSQL & " AND keyword LIKE ?"
			if not(isNull(data_pub)) then 
				if (Application("dbType") = 1) then
					strSQL = strSQL & " AND data_pubblicazione <=?" 	
				else
					strSQL = strSQL & " AND data_pubblicazione <=#?#" 						
				end if
			end if
			if not(isNull(data_del)) then
				if (Application("dbType") = 1) then
					strSQL = strSQL & " AND (data_cancellazione >=? OR data_cancellazione='0000-00-00 00:00:00')" 	
				else
					strSQL = strSQL & " AND data_cancellazione >=#?#" 						
				end if
			end if
			if not(isNull(stato_news)) then strSQL = strSQL & " AND stato_news=?"
			if (hasTarget) then 
				strSQL = strSQL & " AND id IN("					
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
		end if
		
		if not(isNull(order_by)) then
			select Case order_by
			Case 1
				strSQL = strSQL & " ORDER BY titolo ASC"
			Case 2
				strSQL = strSQL & " ORDER BY titolo DESC"
			Case 3
				strSQL = strSQL & " ORDER BY abstract ASC"
			Case 4
				strSQL = strSQL & " ORDER BY abstract DESC"
			Case 5
				strSQL = strSQL & " ORDER BY abstract_2 ASC"
			Case 6
				strSQL = strSQL & " ORDER BY abstract_2 DESC"
			Case 7
				strSQL = strSQL & " ORDER BY abstract_3 ASC"
			Case 8
				strSQL = strSQL & " ORDER BY abstract_3 DESC"
			Case 9
				strSQL = strSQL & " ORDER BY testo ASC"
			Case 10
				strSQL = strSQL & " ORDER BY testo DESC"
			Case 11
				strSQL = strSQL & " ORDER BY data_pubblicazione ASC"
			Case 12
				strSQL = strSQL & " ORDER BY data_pubblicazione DESC"
			Case 13
				strSQL = strSQL & " ORDER BY data_inserimento ASC"
			Case 14
				strSQL = strSQL & " ORDER BY data_inserimento DESC"
			Case 15
				strSQL = strSQL & " ORDER BY keyword ASC"
			Case 16
				strSQL = strSQL & " ORDER BY keyword DESC"
			Case Else
			End Select
		end if

		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"
		
		'response.write(strSQL&"<br>")
		'response.write("time: "&Time()&"<br>")


		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		if (isNull(id) AND isNull(id_utente) AND isNull(titolo) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news)) then
		else
			if not(isNull(id)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
			if not(isNull(id_utente)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
			if not(isNull(titolo)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,titolo)
			if not(isNull(keyword)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&keyword&"%")
			'il passaggio seguente e' da verificare con query secca di test su DB
			if not(isNull(data_pub)) then 
				if (Application("dbType") = 1) then
					data_pub = convertDate(data_pub)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 	
				else
					data_pub = FormatDateTime(data_pub, 2)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 						
				end if
			end if
			if not(isNull(data_del)) then
				if (Application("dbType") = 1) then
					data_del = convertDate(data_del)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del) 	
				else
					data_pub = FormatDateTime(data_del, 2)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del)  						
				end if
			end if
			if not(isNull(stato_news)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,stato_news)
		end if
		Set objRS = objCommand.Execute()


		'response.write("time 2: "&Time()&"<br>")

		if objRS.EOF then
			findNewsSlim = null
		else
			Dim objNews, objListaNews
			Dim objListaTarget, strEditorID, objFiles, objListaFiles, checkTarget
			
			Set objListaNews = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF			
				Dim id_news
				id_news = objRS("id")
				Set objNews = new NewsClass
				objNews.setNewsID(id_news)
				objNews.setTitolo(objRS("titolo"))
				objNews.setAbstract1(objRS("abstract"))
				objNews.setAbstract2(objRS("abstract_2"))
				objNews.setAbstract3(objRS("abstract_3"))
				objNews.setTesto(objRS("testo"))
				objNews.setKeyword(objRS("keyword"))
				objNews.setDataInsNews(objRS("data_inserimento"))
				objNews.setDataPubNews(objRS("data_pubblicazione"))
				objNews.setDataDelNews(objRS("data_cancellazione"))
				objNews.setStato(objRS("stato_news"))		
				objNews.setMetaDescription(objRS("meta_description"))	
				objNews.setMetaKeyword(objRS("meta_keyword"))
				objNews.setPageTitle(objRS("page_title"))
				objNews.setEditorID(objRS("id_utente"))					
				
				objListaNews.add id_news, objNews
				Set objNews = Nothing
				objRS.moveNext()				
			loop
			
			if(bolAddTarget OR bolAddFiles)then								
				Set objNews = new NewsClass				
				Set objFiles = new File4NewsClass			
				for each j in objListaNews
					bolValid = true
					if(bolAddTarget)then	
						Set objListaTarget = objNews.getTargetPerNews(j)								
						if not(isEmpty(objListaTarget)) then
							objListaNews(j).setListaTarget(objListaTarget)
							Set objListaTarget = nothing
						else
							Set objListaTarget = nothing
							call objListaNews.remove(j)
							bolValid = false
						end if
					end if
					
					if(bolValid AND bolAddFiles)then					
						on Error Resume Next				
						Set objListaFiles = objFiles.getFilePerNews(j)			
						if Err.number <> 0 then
							objListaFiles = null
						end if								
						if not(isNull(objListaFiles)) then
							objListaNews(j).setFilePerNews(objListaFiles)
							Set objListaFiles = nothing
						else
							Set objListaFiles = nothing
						end if	
					end if									
				next
				Set objFiles = nothing									
				Set objNews = nothing			
			end if
						
			if (objListaNews.Count > 0) then
				Set findNewsSlim = objListaNews
			else
				findNewsSlim = null			
			end if
			
			Set objListaNews = nothing
			Set objFiles = nothing
		end if
		
		Set objListTarget = nothing
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		'response.write("time 3: "&Time()&"<br>")
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
									

'****** FUNZIONE FIND NEWS CON CACHE, ATTACHMENTS, TARGET MA SENZA FIELD
	Public Function findNewsSlimCached(id, id_utente, titolo, keyword, arrTargetCat, arrTargetLang, data_pub, data_del, stato_news, order_by, bolAddTarget, bolAddFiles)
		findNewsSlimCached = null				

		Dim objDB, strSQL, strSQLTarget, strSQLTmp, objRS, objRSTargetCat, objRSTargetLang, objListTarget, objConn
		Dim hasTarget, doExit, cacheKey
		hasTarget = true
		cacheKey="findc"

		Dim noTargetCat,noTargetlang
		noTargetCat = (isNull(arrTargetCat) OR not(strComp(typename(arrTargetCat), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0))
		
		Set objListTarget = Server.CreateObject("Scripting.Dictionary")
		if (noTargetCat) AND (noTargetlang) then
			hasTarget = false
		elseif (noTargetCat) OR (noTargetlang) then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=026")
		end if
		
		strSQL = "SELECT * FROM news_find"
		if (isNull(id) AND isNull(id_utente) AND isNull(titolo) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news) AND not(hasTarget)) then
			strSQL = "SELECT * FROM news_find"
		else				
			strSQL = strSQL & " WHERE"
			
			Set objBase64 = new Base64Class
			
			if not(isNull(id)) then
				strSQL = strSQL & " AND id=?"
				cacheKey=cacheKey&"-"&id
			end if
			if not(isNull(id_utente)) then 
				strSQL = strSQL & " AND id_utente=?"
				cacheKey=cacheKey&"-"&id_utente
			end if
			if not(isNull(titolo)) then 
				strSQL = strSQL & " AND titolo=?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(titolo)
			end if
			if not(isNull(keyword)) then 
				strSQL = strSQL & " AND keyword LIKE ?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(keyword)
			end if
			if not(isNull(data_pub)) then 
				if (Application("dbType") = 1) then
					strSQL = strSQL & " AND data_pubblicazione <=?" 	
				else
					strSQL = strSQL & " AND data_pubblicazione <=#?#" 						
				end if 
				strSQL = strSQL & " AND keyword LIKE ?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(data_pub)
			end if
			if not(isNull(data_del)) then
				if (Application("dbType") = 1) then
					strSQL = strSQL & " AND (data_cancellazione >=? OR data_cancellazione='0000-00-00 00:00:00')" 	
				else
					strSQL = strSQL & " AND data_cancellazione >=#?#" 						
				end if
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(data_del)
			end if
			if not(isNull(stato_news)) then 
				strSQL = strSQL & " AND stato_news=?"
				cacheKey=cacheKey&"-"&stato_news
			end if
			if (hasTarget) then 
				strSQL = strSQL & " AND id IN("					
				strSQL = strSQL & "SELECT DISTINCT(id_news) FROM target_x_news WHERE id_news IN("
				strSQL = strSQL & "SELECT DISTINCT(id_news) FROM target_x_news WHERE id_target IN("								
				for each idx in arrTargetCat
					strSQL = strSQL &idx&","
					cacheKey=cacheKey&"-"&idx
				next					
				strSQL = strSQL & "))"				
				strSQL = strSQL & " AND id_target IN("	
				for each idy in arrTargetLang
					strSQL = strSQL &idy&","
					cacheKey=cacheKey&"-"&idy
				next					
				strSQL = strSQL & "))"						
				strSQL = Replace(strSQL, ",)", ")", 1, -1, 1)
				strSQL = Trim(strSQL)
			end if
			
			Set objBase64 = nothing
		end if
		
		if not(isNull(order_by)) then
			select Case order_by
			Case 1
				strSQL = strSQL & " ORDER BY titolo ASC"
			Case 2
				strSQL = strSQL & " ORDER BY titolo DESC"
			Case 3
				strSQL = strSQL & " ORDER BY abstract ASC"
			Case 4
				strSQL = strSQL & " ORDER BY abstract DESC"
			Case 5
				strSQL = strSQL & " ORDER BY abstract_2 ASC"
			Case 6
				strSQL = strSQL & " ORDER BY abstract_2 DESC"
			Case 7
				strSQL = strSQL & " ORDER BY abstract_3 ASC"
			Case 8
				strSQL = strSQL & " ORDER BY abstract_3 DESC"
			Case 9
				strSQL = strSQL & " ORDER BY testo ASC"
			Case 10
				strSQL = strSQL & " ORDER BY testo DESC"
			Case 11
				strSQL = strSQL & " ORDER BY data_pubblicazione ASC"
			Case 12
				strSQL = strSQL & " ORDER BY data_pubblicazione DESC"
			Case 13
				strSQL = strSQL & " ORDER BY data_inserimento ASC"
			Case 14
				strSQL = strSQL & " ORDER BY data_inserimento DESC"
			Case 15
				strSQL = strSQL & " ORDER BY keyword ASC"
			Case 16
				strSQL = strSQL & " ORDER BY keyword DESC"
			Case Else
			End Select
			cacheKey=cacheKey&"-"&order_by
		end if

		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"
		
		'response.write(strSQL&"<br>")
		'response.write("time: "&Time()&"<br>")

		cacheKey=Trim(cacheKey)
		'response.write(cacheKey&"<br>")

		'tento il recupero dell'oggetto dalla cache
		on error resume next
		Set ojbCache = new CacheClass
		
		Set cachedObj = ojbCache.getItem(cacheKey)
		
		'response.write("typename(cachedObj): "& typename(cachedObj)&"<br>")
		
		if (Instr(1, typename(cachedObj), "Dictionary", 1) > 0) then
			Set objListaNewsC = Server.CreateObject("Scripting.Dictionary")

			'response.write("cachedObj.count: "& cachedObj.count&"<br>")
		
			for each skey in cachedObj
				Set objNewsC = new NewsClass						
				objNewsC.setNewsID(cachedObj(skey)("id_news"))
				objNewsC.setTitolo(cachedObj(skey)("titolo"))
				objNewsC.setAbstract1(cachedObj(skey)("abstract"))
				objNewsC.setAbstract2(cachedObj(skey)("abstract_2"))
				objNewsC.setAbstract3(cachedObj(skey)("abstract_3"))
				objNewsC.setTesto(cachedObj(skey)("testo"))
				objNewsC.setKeyword(cachedObj(skey)("keyword"))
				objNewsC.setDataInsNews(cachedObj(skey)("data_inserimento"))
				objNewsC.setDataPubNews(cachedObj(skey)("data_pubblicazione"))
				objNewsC.setDataDelNews(cachedObj(skey)("data_cancellazione"))			
				objNewsC.setStato(cachedObj(skey)("stato_news"))	
				objNewsC.setMetaDescription(cachedObj(skey)("meta_description"))	
				objNewsC.setMetaKeyword(cachedObj(skey)("meta_keyword"))
				objNewsC.setPageTitle(cachedObj(skey)("page_title"))
				objNewsC.setEditorID(cachedObj(skey)("id_utente"))
				
				'response.write("titolo:"&objNewsC.getTitolo()&" - id:"&objNewsC.getNewsID()&"<br>")
				
				Set objListaTarget = Server.CreateObject("Scripting.Dictionary")				
				if (Instr(1, typename(cachedObj(skey)("target_list")), "Dictionary", 1) > 0) then
					Set objListaTargetTmp = cachedObj(skey)("target_list")
					for each xt in objListaTargetTmp
						Set objTarget = new Targetclass
						objTarget.setTargetID(xt)
						objTarget.setTargetDescrizione(objListaTargetTmp(xt)("descrizione"))
						objTarget.setTargetType(objListaTargetTmp(xt)("type"))	
						objListaTarget.add xt, objTarget
						'response.write("targetid:"&objTarget.getTargetID()&" - descrizione:"&objTarget.getTargetDescrizione()&"<br>")
						Set objTarget = nothing		
					next				
				
					Set objListaTargetTmp = nothing			
					objNewsC.setListaTarget(objListaTarget)
					'response.write("objNewsC.getListaTarget().count: "& objNewsC.getListaTarget().count&"<br>")
					Set objListaTarget = nothing
				end if			
				
				Set objListaFiles = Server.CreateObject("Scripting.Dictionary")				
				if (Instr(1, typename(cachedObj(skey)("file_list")), "Dictionary", 1) > 0) then
					Set objListaFilesTmp = cachedObj(skey)("file_list")
					for each xf in objListaFilesTmp
						Set objFiles = new File4NewsClass
						objFiles.setFileID(xf)
						objFiles.setFileName(objListaFilesTmp(xf)("filename"))
						objFiles.setFileType(objListaFilesTmp(xf)("content_type"))
						objFiles.setFilePath(objListaFilesTmp(xf)("path"))
						objFiles.setFileDida(objListaFilesTmp(xf)("file_dida"))
						objFiles.setFileTypeLabel(objListaFilesTmp(xf)("file_label"))								
						objListaFiles.add xf, objFiles
						Set objFiles = nothing			
					next
					Set objListaFilesTmp = nothing
				end if
				
				if(objListaFiles.count>0)then
					objNewsC.setFilePerNews(objListaFiles)
				else
					objNewsC.setFilePerNews(null)
				end if
				Set objListaFiles = nothing
								
				'response.write("typename(objNewsC): "& typename(objNewsC)&"<br>")
			
				objListaNewsC.add skey, objNewsC
				Set objNewsC = nothing	
			next	
			'response.write("objListaNewsC.count: "& objListaNewsC.count&"<br>")
			
			Set findNewsSlimCached = objListaNewsC
			Set objListaNewsC = nothing			
		else
			findNewsSlimCached = null
		end if
		
		if Err.number <> 0 then
			findNewsSlimCached = null
			'response.write(Err.number&" - "&Err.description&"<br>")
		end if
		
		'response.write("typename(findNewsSlimCached): "& typename(findNewsSlimCached)&"<br>")

		if not(Instr(1, typename(findNewsSlimCached), "Dictionary", 1) > 0) then
			on error resume next

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
		
			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQL
			
			if (isNull(id) AND isNull(id_utente) AND isNull(titolo) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news)) then
			else
				if not(isNull(id)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
				if not(isNull(id_utente)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
				if not(isNull(titolo)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,titolo)
				if not(isNull(keyword)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&keyword&"%")
				'il passaggio seguente e' da verificare con query secca di test su DB
				if not(isNull(data_pub)) then 
					if (Application("dbType") = 1) then
						data_pub = convertDate(data_pub)
						objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 	
					else
						data_pub = FormatDateTime(data_pub, 2)
						objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 						
					end if
				end if
				if not(isNull(data_del)) then
					if (Application("dbType") = 1) then
						data_del = convertDate(data_del)
						objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del) 	
					else
						data_pub = FormatDateTime(data_del, 2)
						objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del)  						
					end if
				end if
				if not(isNull(stato_news)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,stato_news)
			end if
			Set objRS = objCommand.Execute()

			if objRS.EOF then
				findNewsSlimCached = null
			else
				Dim objNews, objListaNews
				Dim objListaTarget, strEditorID, objFiles, objListaFiles, checkTarget
				
				Set objListaNews = Server.CreateObject("Scripting.Dictionary")
				Set objListaNewsCache = Server.CreateObject("Scripting.Dictionary")
			
				do while not objRS.EOF			
					Dim id_news
					id_news = objRS("id")
					
					Set objNews = new NewsClass
					Set objNewsCache = Server.CreateObject("Scripting.Dictionary")
					
					objNews.setNewsID(id_news)
					objNews.setTitolo(objRS("titolo"))
					objNews.setAbstract1(objRS("abstract"))
					objNews.setAbstract2(objRS("abstract_2"))
					objNews.setAbstract3(objRS("abstract_3"))
					objNews.setTesto(objRS("testo"))
					objNews.setKeyword(objRS("keyword"))
					objNews.setDataInsNews(objRS("data_inserimento"))
					objNews.setDataPubNews(objRS("data_pubblicazione"))
					objNews.setDataDelNews(objRS("data_cancellazione"))
					objNews.setStato(objRS("stato_news"))		
					objNews.setMetaDescription(objRS("meta_description"))	
					objNews.setMetaKeyword(objRS("meta_keyword"))
					objNews.setPageTitle(objRS("page_title"))
					objNews.setEditorID(objRS("id_utente"))					

					objNewsCache.add "id_news", id_news
					objNewsCache.add "titolo", objNews.getTitolo()
					objNewsCache.add "abstract", objNews.getAbstract1()
					objNewsCache.add "abstract_2", objNews.getAbstract2()
					objNewsCache.add "abstract_3", objNews.getAbstract3()
					objNewsCache.add "testo", objNews.getTesto()
					objNewsCache.add "keyword", objNews.getKeyword()
					objNewsCache.add "data_inserimento", objNews.getDataInsNews()
					objNewsCache.add "data_pubblicazione", objNews.getDataPubNews()
					objNewsCache.add "data_cancellazione", objNews.getDataDelNews()
					objNewsCache.add "stato_news", objNews.getStato()
					objNewsCache.add "meta_description", objNews.getMetaDescription()
					objNewsCache.add "meta_keyword", objNews.getMetaKeyword()
					objNewsCache.add "page_title", objNews.getPageTitle()
					objNewsCache.add "id_utente", objNews.getEditorID()	

					objListaNews.add id_news, objNews
					objListaNewsCache.add id_news, objNewsCache
					
					Set objNewsCache = nothing
					Set objNews = Nothing
					objRS.moveNext()				
				loop

				if(bolAddTarget OR bolAddFiles)then								
					Set objNews = new NewsClass				
					Set objFiles = new File4NewsClass			
					for each j in objListaNews
						bolValid = true
						if(bolAddTarget)then	
							Set objListaTarget = objNews.getTargetPerNews(j)	
							'response.write("objListaTarget.count: "& objListaTarget.count&"<br>")							
							if not(isEmpty(objListaTarget)) then
								objListaNews(j).setListaTarget(objListaTarget)
								
								Set objListaTargetCache = Server.CreateObject("Scripting.Dictionary")
								for each xt in objListaTarget
									Set objTargetCache = Server.CreateObject("Scripting.Dictionary")
									objTargetCache.add "id_target", xt
									objTargetCache.add "descrizione", objListaTarget(xt).getTargetDescrizione()
									objTargetCache.add "type", objListaTarget(xt).getTargetType()							
									objListaTargetCache.add xt, objTargetCache	
									'response.write("id_target:"&objTargetCache("id_target")&" - descrizione:"&objTargetCache("descrizione")&"<br>")						
									Set objTargetCache = nothing
								next	
								objListaNewsCache(j).add "target_list", objListaTargetCache
								Set objListaTargetCache = nothing	
							else
								call objListaNews.remove(j)
								call objListaNewsCache.remove(j)
								bolValid = false
							end if
							Set objListaTarget = nothing
						end if
						
						if(bolValid AND bolAddFiles)then					
							on Error Resume Next				
							Set objListaFiles = objFiles.getFilePerNews(j)	
							Set objListaFilesCache = Server.CreateObject("Scripting.Dictionary")			
							if Err.number <> 0 then
								objListaFiles = null
							end if
							
							if not(isNull(objListaFiles)) then
								objListaNews(j).setFilePerNews(objListaFiles)
									
								for each xf in objListaFiles
									Set objFilesCache = Server.CreateObject("Scripting.Dictionary")
									objFilesCache.add "id", xf
									objFilesCache.add "filename", objListaFiles(xf).getFileName()
									objFilesCache.add "content_type", objListaFiles(xf).getFileType()
									objFilesCache.add "path", objListaFiles(xf).getFilePath()
									objFilesCache.add "file_dida", objListaFiles(xf).getFileDida()
									objFilesCache.add "file_label", objListaFiles(xf).getFileTypeLabel()								
									objListaFilesCache.add xf, objFilesCache							
									Set objFilesCache = nothing
								next

								objListaNewsCache(j).add "file_list", objListaFilesCache
								Set objListaFilesCache = nothing
							end if	
							Set objListaFiles = nothing
						end if									
					next
					Set objFiles = nothing									
					Set objNews = nothing			
				end if
							
				if (objListaNews.Count > 0) then
					Set findNewsSlimCached = objListaNews
					call ojbCache.store(cacheKey, objListaNewsCache)
				else
					findNewsSlimCached = null			
				end if
				
				Set objListaNewsCache = nothing
				Set objListaNews = nothing
				Set objFiles = nothing
			end if
			
			Set objListTarget = nothing
					
			Set objRS = Nothing
			Set objCommand = Nothing
			Set objDB = Nothing
			
			if Err.number <> 0 then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
		end if

		Set ojbCache = nothing		
	End Function
									

'****** FUNZIONE FIND NEWS CON CACHE, ATTACHMENTS, TARGET E CON RECUPERO LISTA FIELDS
	Public Function findNewsSlimCachedFields(id, id_utente, titolo, keyword, arrTargetCat, arrTargetLang, data_pub, data_del, stato_news, order_by, bolAddTarget, bolAddFiles, bolAddFields)
		findNewsSlimCachedFields = null				

		Dim objDB, strSQL, strSQLTarget, strSQLTmp, objRS, objRSTargetCat, objRSTargetLang, objListTarget, objConn
		Dim hasTarget, doExit, cacheKey
		hasTarget = true
		cacheKey="findc"

		Dim noTargetCat,noTargetlang
		noTargetCat = (isNull(arrTargetCat) OR not(strComp(typename(arrTargetCat), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0))
		
		Set objListTarget = Server.CreateObject("Scripting.Dictionary")
		if (noTargetCat) AND (noTargetlang) then
			hasTarget = false
		elseif (noTargetCat) OR (noTargetlang) then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=026")
		end if
		
		strSQL = "SELECT * FROM news_find"
		if (isNull(id) AND isNull(id_utente) AND isNull(titolo) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news) AND not(hasTarget)) then
			strSQL = "SELECT * FROM news_find"
		else				
			strSQL = strSQL & " WHERE"
			
			Set objBase64 = new Base64Class
			
			if not(isNull(id)) then
				strSQL = strSQL & " AND id=?"
				cacheKey=cacheKey&"-"&id
			end if
			if not(isNull(id_utente)) then 
				strSQL = strSQL & " AND id_utente=?"
				cacheKey=cacheKey&"-"&id_utente
			end if
			if not(isNull(titolo)) then 
				strSQL = strSQL & " AND titolo=?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(titolo)
			end if
			if not(isNull(keyword)) then 
				strSQL = strSQL & " AND keyword LIKE ?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(keyword)
			end if
			if not(isNull(data_pub)) then 
				if (Application("dbType") = 1) then
					strSQL = strSQL & " AND data_pubblicazione <=?" 	
				else
					strSQL = strSQL & " AND data_pubblicazione <=#?#" 						
				end if 
				strSQL = strSQL & " AND keyword LIKE ?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(data_pub)
			end if
			if not(isNull(data_del)) then
				if (Application("dbType") = 1) then
					strSQL = strSQL & " AND (data_cancellazione >=? OR data_cancellazione='0000-00-00 00:00:00')" 	
				else
					strSQL = strSQL & " AND data_cancellazione >=#?#" 						
				end if
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(data_del)
			end if
			if not(isNull(stato_news)) then 
				strSQL = strSQL & " AND stato_news=?"
				cacheKey=cacheKey&"-"&stato_news
			end if
			if (hasTarget) then 
				strSQL = strSQL & " AND id IN("					
				strSQL = strSQL & "SELECT DISTINCT(id_news) FROM target_x_news WHERE id_news IN("
				strSQL = strSQL & "SELECT DISTINCT(id_news) FROM target_x_news WHERE id_target IN("								
				for each idx in arrTargetCat
					strSQL = strSQL &idx&","
					cacheKey=cacheKey&"-"&idx
				next					
				strSQL = strSQL & "))"				
				strSQL = strSQL & " AND id_target IN("	
				for each idy in arrTargetLang
					strSQL = strSQL &idy&","
					cacheKey=cacheKey&"-"&idy
				next					
				strSQL = strSQL & "))"						
				strSQL = Replace(strSQL, ",)", ")", 1, -1, 1)
				strSQL = Trim(strSQL)
			end if
			
			Set objBase64 = nothing
		end if
		
		if not(isNull(order_by)) then
			select Case order_by
			Case 1
				strSQL = strSQL & " ORDER BY titolo ASC"
			Case 2
				strSQL = strSQL & " ORDER BY titolo DESC"
			Case 3
				strSQL = strSQL & " ORDER BY abstract ASC"
			Case 4
				strSQL = strSQL & " ORDER BY abstract DESC"
			Case 5
				strSQL = strSQL & " ORDER BY abstract_2 ASC"
			Case 6
				strSQL = strSQL & " ORDER BY abstract_2 DESC"
			Case 7
				strSQL = strSQL & " ORDER BY abstract_3 ASC"
			Case 8
				strSQL = strSQL & " ORDER BY abstract_3 DESC"
			Case 9
				strSQL = strSQL & " ORDER BY testo ASC"
			Case 10
				strSQL = strSQL & " ORDER BY testo DESC"
			Case 11
				strSQL = strSQL & " ORDER BY data_pubblicazione ASC"
			Case 12
				strSQL = strSQL & " ORDER BY data_pubblicazione DESC"
			Case 13
				strSQL = strSQL & " ORDER BY data_inserimento ASC"
			Case 14
				strSQL = strSQL & " ORDER BY data_inserimento DESC"
			Case 15
				strSQL = strSQL & " ORDER BY keyword ASC"
			Case 16
				strSQL = strSQL & " ORDER BY keyword DESC"
			Case Else
			End Select
			cacheKey=cacheKey&"-"&order_by
		end if

		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"
		
		'response.write(strSQL&"<br>")
		'response.write("time: "&Time()&"<br>")

		cacheKey=Trim(cacheKey)
		'response.write(cacheKey&"<br>")

		'tento il recupero dell'oggetto dalla cache
		on error resume next
		Set ojbCache = new CacheClass
		
		Set cachedObj = ojbCache.getItem(cacheKey)
		
		'response.write("typename(cachedObj): "& typename(cachedObj)&"<br>")
		
		if (Instr(1, typename(cachedObj), "Dictionary", 1) > 0) then
			Set objListaNewsC = Server.CreateObject("Scripting.Dictionary")

			'response.write("cachedObj.count: "& cachedObj.count&"<br>")
		
			for each skey in cachedObj
				Set objNewsC = new NewsClass						
				objNewsC.setNewsID(cachedObj(skey)("id_news"))
				objNewsC.setTitolo(cachedObj(skey)("titolo"))
				objNewsC.setAbstract1(cachedObj(skey)("abstract"))
				objNewsC.setAbstract2(cachedObj(skey)("abstract_2"))
				objNewsC.setAbstract3(cachedObj(skey)("abstract_3"))
				objNewsC.setTesto(cachedObj(skey)("testo"))
				objNewsC.setKeyword(cachedObj(skey)("keyword"))
				objNewsC.setDataInsNews(cachedObj(skey)("data_inserimento"))
				objNewsC.setDataPubNews(cachedObj(skey)("data_pubblicazione"))
				objNewsC.setDataDelNews(cachedObj(skey)("data_cancellazione"))			
				objNewsC.setStato(cachedObj(skey)("stato_news"))	
				objNewsC.setMetaDescription(cachedObj(skey)("meta_description"))	
				objNewsC.setMetaKeyword(cachedObj(skey)("meta_keyword"))
				objNewsC.setPageTitle(cachedObj(skey)("page_title"))
				objNewsC.setEditorID(cachedObj(skey)("id_utente"))
				
				'response.write("titolo:"&objNewsC.getTitolo()&" - id:"&objNewsC.getNewsID()&"<br>")
				
				Set objListaTarget = Server.CreateObject("Scripting.Dictionary")				
				if (Instr(1, typename(cachedObj(skey)("target_list")), "Dictionary", 1) > 0) then
					Set objListaTargetTmp = cachedObj(skey)("target_list")
					for each xt in objListaTargetTmp
						Set objTarget = new Targetclass
						objTarget.setTargetID(xt)
						objTarget.setTargetDescrizione(objListaTargetTmp(xt)("descrizione"))
						objTarget.setTargetType(objListaTargetTmp(xt)("type"))	
						objListaTarget.add xt, objTarget
						'response.write("targetid:"&objTarget.getTargetID()&" - descrizione:"&objTarget.getTargetDescrizione()&"<br>")
						Set objTarget = nothing		
					next				
				
					Set objListaTargetTmp = nothing			
					objNewsC.setListaTarget(objListaTarget)
					'response.write("objNewsC.getListaTarget().count: "& objNewsC.getListaTarget().count&"<br>")
					Set objListaTarget = nothing
				end if			
				
				Set objListaFiles = Server.CreateObject("Scripting.Dictionary")				
				if (Instr(1, typename(cachedObj(skey)("file_list")), "Dictionary", 1) > 0) then
					Set objListaFilesTmp = cachedObj(skey)("file_list")
					for each xf in objListaFilesTmp
						Set objFiles = new File4NewsClass
						objFiles.setFileID(xf)
						objFiles.setFileName(objListaFilesTmp(xf)("filename"))
						objFiles.setFileType(objListaFilesTmp(xf)("content_type"))
						objFiles.setFilePath(objListaFilesTmp(xf)("path"))
						objFiles.setFileDida(objListaFilesTmp(xf)("file_dida"))
						objFiles.setFileTypeLabel(objListaFilesTmp(xf)("file_label"))								
						objListaFiles.add xf, objFiles
						Set objFiles = nothing			
					next
					Set objListaFilesTmp = nothing
				end if
				
				if(objListaFiles.count>0)then
					objNewsC.setFilePerNews(objListaFiles)
				else
					objNewsC.setFilePerNews(null)
				end if
				Set objListaFiles = nothing			
				
				Set objListaFields = Server.CreateObject("Scripting.Dictionary")				
				if (Instr(1, typename(cachedObj(skey)("field_list")), "Dictionary", 1) > 0) then
					Set objListaFieldsTmp = cachedObj(skey)("field_list")
					for each xf in objListaFieldsTmp
						Set objFields = new ContentFieldClass
						objFields.setFileID(xf)
						objFields.setFileName(objListaFieldsTmp(xf)("filename"))
						objFields.setFileType(objListaFieldsTmp(xf)("content_type"))
						objFields.setFilePath(objListaFieldsTmp(xf)("path"))
						objFields.setFileDida(objListaFieldsTmp(xf)("file_dida"))
						objFields.setFileTypeLabel(objListaFieldsTmp(xf)("file_label"))	

						objFields.setID(xf)
						objFields.setDescription(objListaFieldsTmp(xf)("description"))
						objFields.setIdGroup(objListaFieldsTmp(xf)("id_group"))
						
						Set objGroup = new ContentFieldGroupClass
						Set objCachedGroup = objListaFieldsTmp(xf)("obj_group")
						objGroup.setID(objCachedGroup("id_group"))
						objGroup.setDescription(objCachedGroup("gdesc"))
						objGroup.setOrder(objCachedGroup("gorder"))		
						objFields.setObjGroup(objGroup)	
						Set objCachedGroup = nothing						
						Set objGroup = nothing
						
						objFields.setOrder(objListaFieldsTmp(xf)("order"))	
						objFields.setTypeField(objListaFieldsTmp(xf)("type"))
						objFields.setTypeContent(objListaFieldsTmp(xf)("type_content"))
						objFields.setMaxLenght(objListaFieldsTmp(xf)("max_lenght"))		
						objFields.setRequired(objListaFieldsTmp(xf)("required"))	
						objFields.setEnabled(objListaFieldsTmp(xf)("enabled"))
						objFields.setEditable(objListaFieldsTmp(xf)("editable"))	
						objFields.setidContent(objListaFieldsTmp(xf)("id_news"))
						objFields.setSelValue(objListaFieldsTmp(xf)("value"))					

						objListaFields.add xf, objFields
						Set objFields = nothing			
					next
					Set objListaFieldsTmp = nothing
				end if
				
				if(objListaFields.count>0)then
					objNewsC.setListaFields(objListaFields)
				else
					objNewsC.setListaFields(null)
				end if
				Set objListaFields = nothing

				'response.write("typename(objNewsC): "& typename(objNewsC)&"<br>")
			
				objListaNewsC.add skey, objNewsC
				Set objNewsC = nothing	
			next	
			'response.write("objListaNewsC.count: "& objListaNewsC.count&"<br>")
			
			Set findNewsSlimCachedFields = objListaNewsC
			Set objListaNewsC = nothing			
		else
			findNewsSlimCachedFields = null
		end if
		
		if Err.number <> 0 then
			findNewsSlimCachedFields = null
			'response.write(Err.number&" - "&Err.description&"<br>")
		end if
		
		'response.write("typename(findNewsSlimCachedFields): "& typename(findNewsSlimCachedFields)&"<br>")

		if not(Instr(1, typename(findNewsSlimCachedFields), "Dictionary", 1) > 0) then
			on error resume next

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
		
			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQL
			
			if (isNull(id) AND isNull(id_utente) AND isNull(titolo) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news)) then
			else
				if not(isNull(id)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
				if not(isNull(id_utente)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
				if not(isNull(titolo)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,titolo)
				if not(isNull(keyword)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&keyword&"%")
				'il passaggio seguente e' da verificare con query secca di test su DB
				if not(isNull(data_pub)) then 
					if (Application("dbType") = 1) then
						data_pub = convertDate(data_pub)
						objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 	
					else
						data_pub = FormatDateTime(data_pub, 2)
						objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 						
					end if
				end if
				if not(isNull(data_del)) then
					if (Application("dbType") = 1) then
						data_del = convertDate(data_del)
						objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del) 	
					else
						data_pub = FormatDateTime(data_del, 2)
						objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del)  						
					end if
				end if
				if not(isNull(stato_news)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,stato_news)
			end if
			Set objRS = objCommand.Execute()

			if objRS.EOF then
				findNewsSlimCachedFields = null
			else
				Dim objNews, objListaNews
				Dim objListaTarget, strEditorID, objFiles, objListaFiles, checkTarget
				
				Set objListaNews = Server.CreateObject("Scripting.Dictionary")
				Set objListaNewsCache = Server.CreateObject("Scripting.Dictionary")
			
				do while not objRS.EOF			
					Dim id_news
					id_news = objRS("id")
					
					Set objNews = new NewsClass
					Set objNewsCache = Server.CreateObject("Scripting.Dictionary")
					
					objNews.setNewsID(id_news)
					objNews.setTitolo(objRS("titolo"))
					objNews.setAbstract1(objRS("abstract"))
					objNews.setAbstract2(objRS("abstract_2"))
					objNews.setAbstract3(objRS("abstract_3"))
					objNews.setTesto(objRS("testo"))
					objNews.setKeyword(objRS("keyword"))
					objNews.setDataInsNews(objRS("data_inserimento"))
					objNews.setDataPubNews(objRS("data_pubblicazione"))
					objNews.setDataDelNews(objRS("data_cancellazione"))
					objNews.setStato(objRS("stato_news"))		
					objNews.setMetaDescription(objRS("meta_description"))	
					objNews.setMetaKeyword(objRS("meta_keyword"))
					objNews.setPageTitle(objRS("page_title"))
					objNews.setEditorID(objRS("id_utente"))					

					objNewsCache.add "id_news", id_news
					objNewsCache.add "titolo", objNews.getTitolo()
					objNewsCache.add "abstract", objNews.getAbstract1()
					objNewsCache.add "abstract_2", objNews.getAbstract2()
					objNewsCache.add "abstract_3", objNews.getAbstract3()
					objNewsCache.add "testo", objNews.getTesto()
					objNewsCache.add "keyword", objNews.getKeyword()
					objNewsCache.add "data_inserimento", objNews.getDataInsNews()
					objNewsCache.add "data_pubblicazione", objNews.getDataPubNews()
					objNewsCache.add "data_cancellazione", objNews.getDataDelNews()
					objNewsCache.add "stato_news", objNews.getStato()
					objNewsCache.add "meta_description", objNews.getMetaDescription()
					objNewsCache.add "meta_keyword", objNews.getMetaKeyword()
					objNewsCache.add "page_title", objNews.getPageTitle()
					objNewsCache.add "id_utente", objNews.getEditorID()	

					objListaNews.add id_news, objNews
					objListaNewsCache.add id_news, objNewsCache
					
					Set objNewsCache = nothing
					Set objNews = Nothing
					objRS.moveNext()				
				loop

				if(bolAddTarget OR bolAddFiles OR bolAddFields)then								
					Set objNews = new NewsClass				
					Set objFiles = new File4NewsClass
					Set objContentField = new ContentFieldClass					
					
					if(bolAddFields)then
						on Error Resume Next				
						arrIdNews = Join(objListaNews.keys, ",")							
						Set objFields = objContentField.getListContentField4ContentActiveMultiple(arrIdNews)						

						if Err.number <> 0 then
							objFields = null
						end if						
					end if
					
					for each j in objListaNews
						bolValid = true
						if(bolAddTarget)then	
							Set objListaTarget = objNews.getTargetPerNews(j)	
							'response.write("objListaTarget.count: "& objListaTarget.count&"<br>")							
							if not(isEmpty(objListaTarget)) then
								objListaNews(j).setListaTarget(objListaTarget)
								
								Set objListaTargetCache = Server.CreateObject("Scripting.Dictionary")
								for each xt in objListaTarget
									Set objTargetCache = Server.CreateObject("Scripting.Dictionary")
									objTargetCache.add "id_target", xt
									objTargetCache.add "descrizione", objListaTarget(xt).getTargetDescrizione()
									objTargetCache.add "type", objListaTarget(xt).getTargetType()							
									objListaTargetCache.add xt, objTargetCache	
									'response.write("id_target:"&objTargetCache("id_target")&" - descrizione:"&objTargetCache("descrizione")&"<br>")						
									Set objTargetCache = nothing
								next	
								objListaNewsCache(j).add "target_list", objListaTargetCache
								Set objListaTargetCache = nothing	
							else
								call objListaNews.remove(j)
								call objListaNewsCache.remove(j)
								bolValid = false
							end if
							Set objListaTarget = nothing
						end if
						
						if(bolValid AND bolAddFiles)then					
							on Error Resume Next				
							Set objListaFiles = objFiles.getFilePerNews(j)	
							Set objListaFilesCache = Server.CreateObject("Scripting.Dictionary")			
							if Err.number <> 0 then
								objListaFiles = null
							end if
							
							if not(isNull(objListaFiles)) then
								objListaNews(j).setFilePerNews(objListaFiles)
									
								for each xf in objListaFiles
									Set objFilesCache = Server.CreateObject("Scripting.Dictionary")
									objFilesCache.add "id", xf
									objFilesCache.add "filename", objListaFiles(xf).getFileName()
									objFilesCache.add "content_type", objListaFiles(xf).getFileType()
									objFilesCache.add "path", objListaFiles(xf).getFilePath()
									objFilesCache.add "file_dida", objListaFiles(xf).getFileDida()
									objFilesCache.add "file_label", objListaFiles(xf).getFileTypeLabel()								
									objListaFilesCache.add xf, objFilesCache							
									Set objFilesCache = nothing
								next

								objListaNewsCache(j).add "file_list", objListaFilesCache
								Set objListaFilesCache = nothing
							end if	
							Set objListaFiles = nothing
						end if	
						
						if(bolValid AND bolAddFields)then					
							on Error Resume Next				
							
							Set objListaFields = objFields(j)								
							Set objListaFieldsCache = Server.CreateObject("Scripting.Dictionary")			
							if Err.number <> 0 then
								objListaFields = null
							end if
							
							if not(isNull(objListaFields)) then
								objListaNews(j).setListaFields(objListaFields)
									
								for each xf in objListaFields
									Set objFieldsCache = Server.CreateObject("Scripting.Dictionary")
									objFieldsCache.add "id", xf
									objFieldsCache.add "description", objListaFiles(xf).getDescription()
									objFieldsCache.add "id_group", objListaFiles(xf).getIdGroup()
									
									Set objGroup = Server.CreateObject("Scripting.Dictionary")
									objGroup.add "id_group", objListaFiles(xf).getIdGroup()
									objGroup.add "gdesc", objListaFiles(xf).getObjGroup().getDescription()
									objGroup.add "gorder", objListaFiles(xf).getObjGroup().getOrder()	
									objFieldsCache.add "obj_group", objGroup		
									Set objGroup = nothing
									
									objFieldsCache.add "order", objListaFiles(xf).getOrder()
									objFieldsCache.add "type", objListaFiles(xf).getTypeField()
									objFieldsCache.add "type_content", objListaFiles(xf).getTypeContent()
									objFieldsCache.add "max_lenght", objListaFiles(xf).getMaxLenght()	
									objFieldsCache.add "required", objListaFiles(xf).getRequired()	
									objFieldsCache.add "enabled", objListaFiles(xf).getEnabled()
									objFieldsCache.add "editable", objListaFiles(xf).getEditable()	 
									objFieldsCache.add "id_news", objListaFiles(xf).getidContent()
									objFieldsCache.add "value", objListaFiles(xf).getSelValue()				

									objListaFieldsCache.add xf, objFieldsCache							
									Set objFieldsCache = nothing
								next

								objListaNewsCache(j).add "field_list", objListaFieldsCache
								Set objListaFieldsCache = nothing
							end if	
							Set objListaFields = nothing
						end if									
					next
					Set objContentField = nothing
					Set objFiles = nothing									
					Set objNews = nothing			
				end if
							
				if (objListaNews.Count > 0) then
					Set findNewsSlimCachedFields = objListaNews
					call ojbCache.store(cacheKey, objListaNewsCache)
				else
					findNewsSlimCachedFields = null			
				end if
				
				Set objListaNewsCache = nothing
				Set objListaNews = nothing
				Set objFiles = nothing
			end if
			
			Set objListTarget = nothing
					
			Set objRS = Nothing
			Set objCommand = Nothing
			Set objDB = Nothing
			
			if Err.number <> 0 then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
		end if

		Set ojbCache = nothing		
	End Function

	'****************** funzione di conteggio delle news tramite vari parametri
		
	Public Function countNews(id, id_utente, titolo, keyword, arrTargetCat, arrTargetLang, data_pub, data_del, stato_news)
		on error resume next		
		countNews = 0				
		Dim objDB, strSQL, strSQLTarget, strSQLTmp, objRS, objRSTargetCat, objRSTargetLang, objListTarget, objConn
		Dim hasTarget, doExit
		hasTarget = true
		Dim noTargetCat,noTargetlang
		noTargetCat = (isNull(arrTargetCat) OR not(strComp(typename(arrTargetCat), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0))
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		Set objListTarget = Server.CreateObject("Scripting.Dictionary")
		if (noTargetCat) AND (noTargetlang) then
			hasTarget = false
		elseif (noTargetCat) OR (noTargetlang) then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=026")	
		end if
		
		strSQL = "SELECT count(id) as idc FROM news_find"
		if (isNull(id) AND isNull(id_utente) AND isNull(titolo) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news) AND not(hasTarget)) then
			strSQL = "SELECT count(id) as idc FROM news_find"
		else
			strSQL = strSQL & " WHERE"
			
			if not(isNull(id)) then strSQL = strSQL & " AND id=?"
			if not(isNull(id_utente)) then strSQL = strSQL & " AND id_utente=?"
			if not(isNull(titolo)) then strSQL = strSQL & " AND titolo=?"
			if not(isNull(keyword)) then strSQL = strSQL & " AND keyword LIKE ?"
			if not(isNull(data_pub)) then 
				if (Application("dbType") = 1) then
					strSQL = strSQL & " AND data_pubblicazione <=?" 	
				else
					strSQL = strSQL & " AND data_pubblicazione <=#?#" 						
				end if
			end if
			if not(isNull(data_del)) then
				if (Application("dbType") = 1) then
					strSQL = strSQL & " AND (data_cancellazione >=? OR data_cancellazione='0000-00-00 00:00:00')" 	
				else
					strSQL = strSQL & " AND data_cancellazione >=#?#" 						
				end if
			end if
			if not(isNull(stato_news)) then strSQL = strSQL & " AND stato_news=?"
			if (hasTarget) then 
				strSQL = strSQL & " AND id IN("					
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
		end if

		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		if (isNull(id) AND isNull(id_utente) AND isNull(titolo) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news)) then
		else
			if not(isNull(id)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
			if not(isNull(id_utente)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_utente)
			if not(isNull(titolo)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,titolo)
			if not(isNull(keyword)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&keyword&"%")
			'il passaggio seguente è da verificare con query secca di test su DB
			if not(isNull(data_pub)) then 
				if (Application("dbType") = 1) then
					data_pub = convertDate(data_pub)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 	
				else
					data_pub = FormatDateTime(data_pub, 2)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 						
				end if
			end if
			if not(isNull(data_del)) then
				if (Application("dbType") = 1) then
					data_del = convertDate(data_del)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del) 	
				else
					data_pub = FormatDateTime(data_del, 2)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del)  						
				end if
			end if
			if not(isNull(stato_news)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,stato_news)
		end if
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			countNews = 0
		else				
			countNews = objRS("idc")
		end if
		
		Set objListTarget = nothing
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	
	'****************** funzione di recupero delle news tramite id_news
		
	Public Function findNewsByID(id_news)
		on error resume next
		
		findNewsByID = null	

		Dim objDB, strSQL, strSQLTmp, objRS, objConn

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
					
		strSQL = "SELECT * FROM news_find WHERE id=?;"
		strSQL = Trim(strSQL)
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
		Set objRS = objCommand.Execute()

		if not objRS.EOF then
			Dim objNews
			Dim objListaTarget, strEditorID, objFiles, objListaFile
						
			Set objFiles = new File4NewsClass
				
			Dim this_id_news
			this_id_news = objRS("id")			
			
			Set objNews = new NewsClass			
			
			objNews.setNewsID(this_id_news)
			objNews.setTitolo(objRS("titolo"))
			objNews.setAbstract1(objRS("abstract"))
			objNews.setAbstract2(objRS("abstract_2"))
			objNews.setAbstract3(objRS("abstract_3"))
			objNews.setTesto(objRS("testo"))
			objNews.setKeyword(objRS("keyword"))
			objNews.setDataInsNews(objRS("data_inserimento"))
			objNews.setDataPubNews(objRS("data_pubblicazione"))
			objNews.setDataDelNews(objRS("data_cancellazione"))			
			objNews.setStato(objRS("stato_news"))	
			objNews.setMetaDescription(objRS("meta_description"))	
			objNews.setMetaKeyword(objRS("meta_keyword"))
			objNews.setPageTitle(objRS("page_title"))
			objNews.setEditorID(objRS("id_utente"))				
			
			bolValid = true
					
			Set objListaTarget = objNews.getTargetPerNews(this_id_news)
			if not(isEmpty(objListaTarget)) then
				objNews.setListaTarget(objListaTarget)
				Set objListaTarget = nothing
			else
				findNewsByID = null
				bolValid = false
			end if			

			if(bolValid)then
				Set objListaFiles = objFiles.getFilePerNews(this_id_news)				
				if not(isEmpty(objListaFiles)) then
					objNews.setFilePerNews(objListaFiles)
					Set objListaFiles = nothing
				else
					objNews.setFilePerNews(null)
					Set objListaFiles = nothing
				end if
			
				Set findNewsByID = objNews
			end if
			
			Set objNews = nothing
			Set objFiles = nothing
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	
	'****************** funzione di recupero delle news tramite id_news cached
		
	Public Function findNewsByIDCached(id_news)		
		findNewsByIDCached = null	
		
		'tento il recupero dell'oggetto dalla cache
		on error resume next
		Set ojbCache = new CacheClass
		
		'Set cachedObj = caching.item("content-"&id_news)
		Set cachedObj = ojbCache.getItem("content-"&id_news)
		
		if (Instr(1, typename(cachedObj), "Dictionary", 1) > 0) then
			Set objNewsC = new NewsClass						
			objNewsC.setNewsID(cachedObj("id_news"))
			objNewsC.setTitolo(cachedObj("titolo"))
			objNewsC.setAbstract1(cachedObj("abstract"))
			objNewsC.setAbstract2(cachedObj("abstract_2"))
			objNewsC.setAbstract3(cachedObj("abstract_3"))
			objNewsC.setTesto(cachedObj("testo"))
			objNewsC.setKeyword(cachedObj("keyword"))
			objNewsC.setDataInsNews(cachedObj("data_inserimento"))
			objNewsC.setDataPubNews(cachedObj("data_pubblicazione"))
			objNewsC.setDataDelNews(cachedObj("data_cancellazione"))			
			objNewsC.setStato(cachedObj("stato_news"))	
			objNewsC.setMetaDescription(cachedObj("meta_description"))	
			objNewsC.setMetaKeyword(cachedObj("meta_keyword"))
			objNewsC.setPageTitle(cachedObj("page_title"))
			objNewsC.setEditorID(cachedObj("id_utente"))
			
			Set objListaTarget = Server.CreateObject("Scripting.Dictionary")
			Set objListaTargetTmp = cachedObj("target_list")
			for each xt in objListaTargetTmp
				Set objTarget = new Targetclass
				objTarget.setTargetID(xt)
				objTarget.setTargetDescrizione(objListaTargetTmp(xt)("descrizione"))
				objTarget.setTargetType(objListaTargetTmp(xt)("type"))	
				objListaTarget.add xt, objTarget
				Set objTarget = nothing		
			next
			Set objListaTargetTmp = nothing			
			objNewsC.setListaTarget(objListaTarget)
			Set objListaTarget = nothing
			
			Set objListaFiles = Server.CreateObject("Scripting.Dictionary")
			Set objListaFilesTmp = cachedObj("file_list")
			for each xf in objListaFilesTmp
				Set objFiles = new File4NewsClass
				objFiles.setFileID(xf)
				objFiles.setFileName(objListaFilesTmp(xf)("filename"))
				objFiles.setFileType(objListaFilesTmp(xf)("content_type"))
				objFiles.setFilePath(objListaFilesTmp(xf)("path"))
				objFiles.setFileDida(objListaFilesTmp(xf)("file_dida"))
				objFiles.setFileTypeLabel(objListaFilesTmp(xf)("file_label"))								
				objListaFiles.add xf, objFiles
				Set objFiles = nothing			
			next
			Set objListaFilesTmp = nothing
			if(objListaFiles.count>0)then
				objNewsC.setFilePerNews(objListaFiles)
			else
				objNewsC.setFilePerNews(null)
			end if
			Set objListaFiles = nothing
			
			Set findNewsByIDCached = objNewsC
			Set objNewsC = nothing			
		else
			findNewsByIDCached = null
		end if
		
		if Err.number <> 0 then
			findNewsByIDCached = null
			'response.write(Err.number&" - "&Err.description&"<br>")
		end if
		

		if not(Instr(1, typename(findNewsByIDCached), "NewsClass", 1) > 0) then
			on error resume next
			Dim objDB, strSQL, strSQLTmp, objRS, objConn

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
						
			strSQL = "SELECT * FROM news_find WHERE id=?;"
			strSQL = Trim(strSQL)
			
			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQL
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
			Set objRS = objCommand.Execute()

			if not objRS.EOF then
				Dim objNews
				Dim objListaTarget, strEditorID, objFiles, objListaFile
							
				Set objFiles = new File4NewsClass
					
				Dim this_id_news
				this_id_news = objRS("id")			
				
				Set objNews = new NewsClass
				Set objNewsCache = Server.CreateObject("Scripting.Dictionary")				
				
				objNews.setNewsID(this_id_news)
				objNews.setTitolo(objRS("titolo"))
				objNews.setAbstract1(objRS("abstract"))
				objNews.setAbstract2(objRS("abstract_2"))
				objNews.setAbstract3(objRS("abstract_3"))
				objNews.setTesto(objRS("testo"))
				objNews.setKeyword(objRS("keyword"))
				objNews.setDataInsNews(objRS("data_inserimento"))
				objNews.setDataPubNews(objRS("data_pubblicazione"))
				objNews.setDataDelNews(objRS("data_cancellazione"))			
				objNews.setStato(objRS("stato_news"))	
				objNews.setMetaDescription(objRS("meta_description"))	
				objNews.setMetaKeyword(objRS("meta_keyword"))
				objNews.setPageTitle(objRS("page_title"))
				objNews.setEditorID(objRS("id_utente"))	

				objNewsCache.add "id_news", this_id_news
				objNewsCache.add "titolo", objNews.getTitolo()
				objNewsCache.add "abstract", objNews.getAbstract1()
				objNewsCache.add "abstract_2", objNews.getAbstract2()
				objNewsCache.add "abstract_3", objNews.getAbstract3()
				objNewsCache.add "testo", objNews.getTesto()
				objNewsCache.add "keyword", objNews.getKeyword()
				objNewsCache.add "data_inserimento", objNews.getDataInsNews()
				objNewsCache.add "data_pubblicazione", objNews.getDataPubNews()
				objNewsCache.add "data_cancellazione", objNews.getDataDelNews()
				objNewsCache.add "stato_news", objNews.getStato()
				objNewsCache.add "meta_description", objNews.getMetaDescription()
				objNewsCache.add "meta_keyword", objNews.getMetaKeyword()
				objNewsCache.add "page_title", objNews.getPageTitle()
				objNewsCache.add "id_utente", objNews.getEditorID()	
				
				bolValid = true
				
				Set objListaTarget = objNews.getTargetPerNews(this_id_news)
				if not(isEmpty(objListaTarget)) then
					objNews.setListaTarget(objListaTarget)
					
					Set objListaTargetCache = Server.CreateObject("Scripting.Dictionary")
					for each xt in objListaTarget
						Set objTargetCache = Server.CreateObject("Scripting.Dictionary")
						objTargetCache.add "id_target", xt
						objTargetCache.add "descrizione", objListaTarget(xt).getTargetDescrizione()
						objTargetCache.add "type", objListaTarget(xt).getTargetType()							
						objListaTargetCache.add xt, objTargetCache							
						Set objTargetCache = nothing
					next	
					objNewsCache.add "target_list", objListaTargetCache
					Set objListaTargetCache = nothing				
				
					Set objListaTarget = nothing
				else
					Set objListaTarget = nothing
					findNewsByIDCached = null
					bolValid = false
				end if			

				if(bolValid)then
					Set objListaFiles = objFiles.getFilePerNews(this_id_news)
					Set objListaFilesCache = Server.CreateObject("Scripting.Dictionary")								
					if not(isEmpty(objListaFiles)) then
						objNews.setFilePerNews(objListaFiles)
						
						for each xf in objListaFiles
							Set objFilesCache = Server.CreateObject("Scripting.Dictionary")
							objFilesCache.add "id", xf
							objFilesCache.add "filename", objListaFiles(xf).getFileName()
							objFilesCache.add "content_type", objListaFiles(xf).getFileType()
							objFilesCache.add "path", objListaFiles(xf).getFilePath()
							objFilesCache.add "file_dida", objListaFiles(xf).getFileDida()
							objFilesCache.add "file_label", objListaFiles(xf).getFileTypeLabel()								
							objListaFilesCache.add xf, objFilesCache							
							Set objFilesCache = nothing
						next
						
						Set objListaFiles = nothing
					else
						objNews.setFilePerNews(null)
						Set objListaFiles = nothing
					end if
					objNewsCache.add "file_list", objListaFilesCache
					Set objListaFilesCache = nothing					
					
					Set findNewsByIDCached = objNews

					call ojbCache.store("content-"&this_id_news, objNewsCache)					
					Set objNewsCache = nothing					
				end if
				
				Set objNews = nothing
				Set objFiles = nothing
			end if
					
			Set objRS = Nothing
			Set objCommand = Nothing
			Set objDB = Nothing
			
			if Err.number <> 0 then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if
		end if

		Set ojbCache = nothing
	End Function
	
	Public Sub insertTargetXNews(id_target, id_news, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO target_x_news(id_target, id_news) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
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
		
	Public Sub insertTargetXNewsNoTransaction(id_target, id_news)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO target_x_news(id_target, id_news) VALUES("
		strSQL = strSQL & "?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteTargetXNews(id_news, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM target_x_news WHERE id_news=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
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
		
	Public Sub deleteTargetXNewsNoTransaction(id_news)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM target_x_news WHERE id_news=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Function getTargetPerNews(id_news)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getTargetPerNews = null		
		strSQL = "SELECT target_x_news.id_target, target.descrizione, target.type FROM target INNER JOIN target_x_news ON target.id = target_x_news.id_target WHERE target_x_news.id_news=? ORDER BY target_x_news.id_target;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
		Set objRS = objCommand.Execute()	

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objTarget = new Targetclass
				strID = objRS("id_target")
				objTarget.setTargetID(strID)
				objTarget.setTargetDescrizione(objRS("descrizione"))
				objTarget.setTargetType(objRS("type"))	
				objDict.add strID, objTarget
				Set objTarget = nothing
				objRS.moveNext()
			loop
							
			Set getTargetPerNews = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
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
%>