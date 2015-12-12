<%

Class LanguageClass
	Private id
	Private descrizione
	Private label
	Private lang_active
	Private sub_domain_active
	Private url_subdomain
	
	Public Function getLanguageID()
		getLanguageID = id
	End Function
	
	Public Sub setLanguageID(strID)
		id = strID
	End Sub
	
	Public Function getLanguageDescrizione()
		getLanguageDescrizione = descrizione
	End Function
	
	Public Sub setLanguageDescrizione(strDesc)
		descrizione = strDesc
	End Sub
	
	Public Function getLabelDescrizione()
		getLabelDescrizione = label
	End Function
	
	Public Sub setLabelDescrizione(strDesc)
		label = strDesc
	End Sub
	
	Public Sub setLangActive(strLangActive)
		lang_active = strLangActive
	End Sub
	
	Public Function isLangActive()
		isLangActive = lang_active
	End Function
	
	Public Sub setSubDomainActive(strSubDomainActive)
		sub_domain_active = strSubDomainActive
	End Sub
	
	Public Function isSubDomainActive()
		isSubDomainActive = sub_domain_active
	End Function
	
	Public Function getURLSubDomain()
		getURLSubDomain = url_subdomain
	End Function
	
	Public Sub setURLSubDomain(strURLSubDomain)
		url_subdomain = strURLSubDomain
	End Sub
	
	Public Function getListaLangDisponibili()
		getListaLangDisponibili = null
		
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict	
		strSQL = "SELECT * FROM language_disponibili;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("keyword")
				strDesc = objRS("description")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaLangDisponibili = objDict			
			Set objDict = nothing				
		end if		
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function	
		
	Public Function getListaLanguage()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaLanguage = null		
		strSQL = "SELECT * FROM language;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strLabel = objRS("label")
				strDesc = objRS("descrizione")
				bolLangActive = objRS("lang_active")
				bolSubDomainActive = objRS("subdomain_active")
				strURLSubDomain = objRS("url_subdomain")
						
				Set objLanguageTmp = new LanguageClass
				objLanguageTmp.setLanguageID(strID)
				objLanguageTmp.setLabelDescrizione(strLabel)	
				objLanguageTmp.setLangActive(bolLangActive)	
				objLanguageTmp.setLanguageDescrizione(strDesc)	
				objLanguageTmp.setSubDomainActive(bolSubDomainActive)		
				objLanguageTmp.setURLSubDomain(strURLSubDomain)									
				objDict.add strID, objLanguageTmp
				Set objLanguageTmp = Nothing
				objRS.moveNext()
			loop
							
			Set getListaLanguage = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListaLanguageByDesc()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaLanguageByDesc = null		
		'strSQL = "SELECT * FROM language ORDER BY descrizione DESC;"
		strSQL = "SELECT * FROM language ORDER BY descrizione ASC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("descrizione")
				strDesc = objRS("label")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaLanguageByDesc = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function getListaLanguageByDescExt()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaLanguageByDescExt = null		
		strSQL = "SELECT * FROM language ORDER BY descrizione ASC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strLabel = objRS("label")
				strDesc = objRS("descrizione")
				bolLangActive = objRS("lang_active")
				bolSubDomainActive = objRS("subdomain_active")
				strURLSubDomain = objRS("url_subdomain")
						
				Set objLanguageTmp = new LanguageClass
				objLanguageTmp.setLanguageID(strID)
				objLanguageTmp.setLabelDescrizione(strLabel)	
				objLanguageTmp.setLangActive(bolLangActive)	
				objLanguageTmp.setLanguageDescrizione(strDesc)	
				objLanguageTmp.setSubDomainActive(bolSubDomainActive)		
				objLanguageTmp.setURLSubDomain(strURLSubDomain)									
				objDict.add strID, objLanguageTmp
				Set objLanguageTmp = Nothing
				objRS.moveNext()
			loop
							
			Set getListaLanguageByDescExt = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
			
	Public Function findLanguage(id_lang)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findLanguage = null		
		strSQL = "SELECT * FROM language WHERE id=?;"
		
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
			Dim objLanguage
			Set objLanguage = new LangugeClass
			objLanguage.setLanguageID(objRS("id"))
			objLanguage.setLanguageDescrizione(objRS("label"))	
			objLanguage.setLangActive(objRS("lang_active"))
			objLanguage.setSubDomainActive(objRS("subdomain_active"))					
			objLanguageTmp.setURLSubDomain(objRS("url_subdomain"))			
			Set findLanguage = objLanguage
			Set objLanguage = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Function getCurrURLSubDomainByLangCode(lang_desc)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getCurrURLSubDomainByLangCode = ""		
		strSQL = "SELECT url_subdomain FROM language WHERE descrizione=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,lang_desc)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			getCurrURLSubDomainByLangCode = objRS("url_subdomain")

		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function isLanguageSelected(strDesc)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		isLanguageSelected = false		
		strSQL = "SELECT * FROM language WHERE descrizione=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDesc)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			isLanguageSelected = true				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function isLanguageSelectedSubDomainActive(strDesc)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		isLanguageSelectedSubDomainActive = false		
		strSQL = "SELECT * FROM language WHERE descrizione=? AND subdomain_active=1;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDesc)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then
			isLanguageSelectedSubDomainActive = true				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
			
	Public Function countActiveLanguage()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		countActiveLanguage = 0		
		strSQL = "SELECT count(id) as idc FROM language WHERE lang_active=1;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			countActiveLanguage = Cint(objRS("idc"))			
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
				
	Public Sub insertLanguage(strDescrizione, strLabel, langActive, subDomainActive, strURLsubdomain)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO language(descrizione, label, lang_active, subdomain_active,url_subdomain) VALUES("
		strSQL = strSQL & "?,?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strLabel)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,langActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,subDomainActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strURLsubdomain)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteLanguage(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM language WHERE id=?;" 

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
				
	Public Sub updateActiveLanguage(id, subLangActive)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE language SET lang_active=? WHERE id =?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,subLangActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

'*********************************************************************************************************
'*********************************************************************************************************
'*************************** CODICE PER IL MULTILINGUA ***************************************************	


	Private langCode
	Private defLangCode
	Private langElements	
	
	Public Sub setLangCode(strCode)
		langCode = strCode			
	end Sub
	
	Public Function getLangCode()
		getLangCode = langCode	
	end Function
	
	Public Sub setDefaultLangCode(strDefCode)
		defLangCode = strDefCode			
	end Sub
	
	Public Function getDefaultLangCode()
		getDefaultLangCode = defLangCode	
	end Function
	
	Public Sub setLangElements(strElements)
		Set langElements = strElements			
	end Sub
	
	Public Function getLangElements()
		Set getLangElements = langElements	
	end Function
					
	Public Sub insertMultiLanguage(strKeyword, strLangCode, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO multi_languages(keyword,lang_code,value) VALUES(?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,strKeyword)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,strLangCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strValue)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
				
	Public Sub modifyMultiLanguage(strKeyword, strLangCode, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, strSQLSelect, strSQL2, foundKeyword
		
		foundKeyword = false
		
		strSQLSelect = "SELECT * FROM multi_languages WHERE keyword =? AND lang_code=?;"
		strSQL = "UPDATE multi_languages SET keyword=?,lang_code=?,value=? WHERE keyword =? AND lang_code=?;"
		strSQL2 = "INSERT INTO multi_languages(keyword,lang_code,value) VALUES(?,?,?);"
		
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
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,strKeyword)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,strLangCode)
		objCommand2.CommandText = strSQL
		objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,150,strKeyword)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,10,strLangCode)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,201,1,-1,strValue)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,150,strKeyword)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,10,strLangCode)
		objCommand3.CommandText = strSQL2
		objCommand3.Parameters.Append objCommand3.CreateParameter(,200,1,150,strKeyword)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,200,1,10,strLangCode)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,201,1,-1,strValue)
		
		objConn.BeginTrans		

		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then
			objCommand2.Execute()
		else
			objCommand3.Execute()		
		end if

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
				
	Public Function updateKeywordMultilanguage(id_multi_language, strKeyword)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, strSQLSelect, strSQL2, foundKeyword
		
		updateKeywordMultilanguage = false
		
		strSQLSelect = "SELECT keyword FROM multi_languages WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLSelect
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_multi_language)
		
		objConn.BeginTrans

		'** recupero la keyword attuale in base all'id e imposto la query di update per le altre lingue		
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then
			strSQL2 = "UPDATE multi_languages SET keyword=? WHERE keyword =?;"
			Dim objCommand2
			Set objCommand2 = Server.CreateObject("ADODB.Command")
			objCommand2.ActiveConnection = objConn
			objCommand2.CommandText = strSQL2
			objCommand2.CommandType=1
			objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,150,strKeyword)
			objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,150,objRS("keyword"))		
			objCommand2.Execute()
			updateKeywordMultilanguage = true
		end if
		Set objRS = Nothing
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
	End Function
			
	Public Sub deleteMultiLanguage(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, strSQLSelect, foundKeyword, strSQL2
		
		foundKeyword = false
		
		strSQLSelect = "SELECT keyword FROM multi_languages WHERE id=?;"
		strSQL = "DELETE FROM multi_languages WHERE id=?;" 

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
		objCommand.CommandText = strSQLSelect
		objCommand2.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		
		objConn.BeginTrans

		'** recupero la keyword attuale in base all'id e imposto la query di update per le altre lingue
		Set objRS = objCommand.Execute()
		if not(objRS.EOF) then
			strSQL2 = "DELETE FROM multi_languages  WHERE keyword =?;"
			objCommand3.CommandText = strSQL2
			objCommand3.Parameters.Append objCommand3.CreateParameter(,200,1,150,objRS("keyword"))
			foundKeyword = true
		end if


		objCommand2.Execute()	
		if(foundKeyword) then
			objCommand3.Execute()
		end if		
		Set objRS = Nothing
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

	Public Function searchDistinctKeyList(strKey)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		searchDistinctKeyList = null		
		strSQL = "SELECT DISTINCT keyword,id FROM multi_languages WHERE keyword LIKE ?  OR value LIKE ? ORDER BY keyword;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,"%"&strKey&"%")
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&strKey&"%")
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			tmpKey = ""
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("keyword")	
				if(tmpKey<>strDesc) then		
					objDict.add strDesc, strID
					tmpKey = strDesc
				end if
				objRS.moveNext()
			loop
						
			Set searchDistinctKeyList = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function searchFilteredListElementsByKey(strKey)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		searchFilteredListElementsByKey = null		
		strSQL = "SELECT * FROM multi_languages WHERE keyword =? ORDER BY lang_code ASC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,strKey)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strKey = objRS("keyword")
				strLangcode = objRS("lang_code")	
				strLabel = objRS("value")		
				objDict.add strLangcode&"-"&strKey, strLabel
				objRS.moveNext()
			loop
						
			Set searchFilteredListElementsByKey = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function searchListaElementsByLang(langCode,strKey)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		searchListaElementsByLang = null		
		strSQL = "SELECT id, value, keyword FROM multi_languages WHERE lang_code=? AND (keyword LIKE ?  OR value LIKE ?) ORDER BY keyword;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,langCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,"%"&strKey&"%")
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&strKey&"%")
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("keyword")
				strLabel = objRS("value")					
				Set objLanguageTmp = new LanguageClass
				objLanguageTmp.setLanguageID(strID)	
				objLanguageTmp.setLanguageDescrizione(strDesc)			
				objLanguageTmp.setLabelDescrizione(strLabel)						
				objDict.add strID, objLanguageTmp
				Set objLanguageTmp = Nothing
				objRS.moveNext()
			loop
						
			Set searchListaElementsByLang = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getListaElementsByLang(langCode)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaElementsByLang = null		
		strSQL = "SELECT value, keyword FROM multi_languages WHERE lang_code=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,langCode)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("keyword")
				strDesc = objRS("value")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaElementsByLang = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getListaElementsByLangAndKey(langCode, strKey)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaElementsByLangAndKey = null		
		strSQL = "SELECT value, keyword FROM multi_languages WHERE lang_code=? AND keyword LIKE ? ORDER BY value;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,langCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,"%"&strKey&"%")
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("keyword")
				strDesc = objRS("value")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaElementsByLangAndKey = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getTranslated(strKeyword)
		on error resume next
		Dim objDB, strSQL, strSQL2, objRS, objRS2, objConn, strTmp, foundKey, foundKeyOnMemory
		getTranslated = ""		
		
		foundKey = false	
		notTriedFoundKeyOnMemory = true			
		
		if not(isNull(langElements)) AND (Instr(1, typename(langElements), "dictionary", 1) > 0) then
			if(langElements.Exists(strKeyword)) AND not(langElements.item(strKeyword)="") then
				getTranslated = langElements.item(strKeyword)
				foundKey = true
			else			
				notTriedFoundKeyOnMemory = false
			end if
		end if
		
		if not(foundKey) then
			strSQL = "SELECT value FROM multi_languages WHERE lang_code=? AND keyword=?;"
			strSQL2 = "SELECT value FROM multi_languages WHERE lang_code=? AND keyword=?;"			
					
			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()		

			Dim objCommand, objCommand2
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand2 = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand2.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand2.CommandType=1
			objCommand.CommandText = strSQL
			objCommand2.CommandText = strSQL2
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,langCode)
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,150,strKeyword)
			objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,10,defLangCode)
			objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,150,strKeyword)

			if (notTriedFoundKeyOnMemory) then
				Set objRS = objCommand.Execute()
				'Set objRS = objConn.Execute(strSQL)				
				if not(objRS.EOF) then
					strTmp = objRS("value")
					if not(strComp(strTmp, "", 1) = 0) then
						getTranslated = strTmp
						foundKey = true
					end if		
				end if
				Set objRS = Nothing
			end if
			
			if not(foundKey) then
				Set objRS2 = objCommand2.Execute()
				'Set objRS2 = objConn.Execute(strSQL2)
				if not(objRS2.EOF) then
					strTmp = objRS2("value")
					if not(strComp(strTmp, "", 1) = 0) then
						getTranslated = strTmp
					end if				
				end if		
				Set objRS2 = Nothing		
			end if
			
			Set objCommand = Nothing
			Set objCommand2 = Nothing
			Set objDB = Nothing
		end if
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

End Class
%>