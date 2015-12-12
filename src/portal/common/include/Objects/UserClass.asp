<%
Class UserClass
	Private idUser
	Private usrName
	Private pwd
	Private email
	Private ruolo
	Private privacy
	Private newsletter
	private userActive
	private friendActive
	private sconto
	private adminComments
	private insertDate
	private modifyDate
	private confirmationCode
	private publicProfile
	private userGroup
	private automaticUser

	'****************** GET FUNCTIONS ***************
	
	Public Function getUserID()
		getUserID = idUser
	End Function
		
	Public Function getUserName()
		getUserName = usrName
	End Function
	
	Public Function getPassword()
		getPassword = pwd
	End Function
	
	Public Function getEmail()
		getEmail = email
	End Function
	
	Public Function getRuolo()
		getRuolo = ruolo
	End Function

	Public Function getPrivacy()
		getPrivacy = Cbool(privacy)
	End Function

	Public Function getNewsletter()
		getNewsletter = Cbool(newsletter)
	End Function
	
	Public Function getUserActive()
		getUserActive = Cint(userActive)
	End Function
	
	Public Function getFriendActive()
		getFriendActive = Cint(friendActive)
	End Function
	
	Public Function getSconto()
		getSconto = sconto
	End Function
	
	Public Function getAdminComments()
		getAdminComments = adminComments
	End Function
	
	Public Function getInsertDate()
		getInsertDate = insertDate
	End Function
	
	Public Function getModifyDate()
		getModifyDate = modifyDate
	End Function

	Public Function getConfirmationCode()
		getConfirmationCode = confirmationCode
	End Function

	Public Function getPublic()
		getPublic = publicProfile
	End Function

	Public Function getGroup()
		getGroup = userGroup
	End Function

	Public Function getAutomaticUser()
		getAutomaticUser = automaticUser
	End Function
	
	
	'****************** SET FUNCTIONS ***************
				
	Public Sub setUserID(strID)
		idUser = strID
	End Sub
		
	Public Sub setUserName(strUserName)
		usrName = strUserName
	End Sub
	
	Public Sub setPassword(strPwd)
		pwd = strPwd
	End Sub
	
	Public Sub setEmail(strEmail)
		email = strEmail
	End Sub
	
	Public Sub setRuolo(ruolo_)
		ruolo = ruolo_
	End Sub

	Public Sub setPrivacy(privacy_)
		privacy = privacy_
	End Sub

	Public Sub setNewsletter(newsletter_)
		newsletter = newsletter_
	End Sub	
	
	Public Sub setUserActive(strUserActive)
		userActive = strUserActive
	End Sub
	
	Public Sub setFriendActive(strFriendActive)
		friendActive = strFriendActive
	End Sub
	
	Public Sub setSconto(numSconto)
		sconto = numSconto
	End Sub
	
	Public Sub setAdminComments(strAdminComments)
		adminComments = strAdminComments
	End Sub
	
	Public Sub setInsertDate(dateInsertDate)
		insertDate = dateInsertDate
	End Sub
	
	Public Sub setModifyDate(dateModifyDate)
		modifyDate = dateModifyDate
	End Sub

	Public Sub setConfirmationCode(numConfirmationCode)
		confirmationCode = numConfirmationCode
	End Sub	

	Public Sub setPublic(strPublic)
		publicProfile = strPublic
	End Sub

	Public Sub setGroup(strGroup)
		userGroup = strGroup
	End Sub

	Public Sub setAutomaticUser(bolAutomaticUser)
		automaticUser = bolAutomaticUser
	End Sub

		
	Public Function findUtente(userNameOrMail, ruolo, user_active, publicProfile, automatic_user, order_by)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objUtente
		findUtente = null  

		strSQL = "SELECT * FROM utenti"
		if (isNull(userNameOrMail) AND isNull(ruolo) AND isNull(user_active) AND isNull(publicProfile) AND isNull(automatic_user)) then
			strSQL = "SELECT * FROM utenti"
		else
			strSQL = strSQL & " WHERE"
			if not(isNull(userNameOrMail)) then strSQL = strSQL & " AND (username =? OR email=?)"
			if not(isNull(ruolo)) then strSQL = strSQL & " AND ruolo IN("&ruolo&")"
			if not(isNull(user_active)) then strSQL = strSQL & " AND utenteAttivo=?"
			if not(isNull(publicProfile)) then strSQL = strSQL & " AND public=?"
			if not(isNull(automatic_user)) then strSQL = strSQL & " AND automatic_user=?"
		end if
		
		if not(isNull(order_by)) then
			select Case order_by
			Case 1
				strSQL = strSQL & " ORDER BY username ASC"
			Case 2
				strSQL = strSQL & " ORDER BY username DESC"
			Case 3
				strSQL = strSQL & " ORDER BY ruolo ASC"
			Case 4
				strSQL = strSQL & " ORDER BY ruolo DESC"
			Case 5
				strSQL = strSQL & " ORDER BY utenteAttivo ASC"
			Case 6
				strSQL = strSQL & " ORDER BY utenteAttivo DESC"
			Case 7
				strSQL = strSQL & " ORDER BY public ASC"
			Case 8
				strSQL = strSQL & " ORDER BY public DESC"
			Case Else
				strSQL = strSQL & " ORDER BY username ASC"
			End Select
		end if

		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = strSQL & ";"
		strSQL = Trim(strSQL)
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		if (isNull(userNameOrMail) AND isNull(user_active) AND isNull(publicProfile) AND isNull(automatic_user)) then
		else
			if not(isNull(userNameOrMail)) then 
				objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,userNameOrMail)
				objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,userNameOrMail)
			end if
			if not(isNull(user_active)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,user_active)
			if not(isNull(publicProfile)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,publicProfile)
			if not(isNull(automatic_user)) then objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,automatic_user)
		end if
		Set objRS = objCommand.Execute()		

		objDict = null
		if not(objRS.EOF) then   
			Set objDict = Server.CreateObject("Scripting.Dictionary")   
			do while not objRS.EOF
				Set objUtente = new UserClass
				strID = objRS("id")
				
				objUtente.setUserID(objRS("id"))    
				objUtente.setUserName(objRS("username"))
				objUtente.setPassword(objRS("pwd"))
				objUtente.setEmail(objRS("email"))
				objUtente.setPrivacy(objRS("privacy"))
				objUtente.setNewsletter(objRS("newsletter"))
				objUtente.setRuolo(objRS("ruolo"))
				objUtente.setUserActive(objRS("utenteAttivo"))
				objUtente.setSconto(objRS("sconto"))
				objUtente.setAdminComments(objRS("adminComments"))
				objUtente.setInsertDate(objRS("insertDate"))
				objUtente.setModifyDate(objRS("modifyDate"))
				objUtente.setPublic(objRS("public"))
				objUtente.setGroup(objRS("user_group"))
				objUtente.setAutomaticUser(objRS("automatic_user"))
				
				
				objDict.add strID, objUtente	
				Set objUtente = nothing
				objRS.moveNext()
			loop
						
			if(not(isNull(objDict))) then
				Set findUtente = objDict  
				Set objDict = nothing  
			else
				findUtente = null
			end if   
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
		
	Public Function getListaUtenti()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objUtente
		getListaUtenti = null  
		strSQL = "SELECT * FROM utenti ORDER BY username;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then   
		Set objDict = Server.CreateObject("Scripting.Dictionary")   
		do while not objRS.EOF
		Set objUtente = new UserClass
		strID = objRS("id")
		
		objUtente.setUserID(strID)    
		objUtente.setUserName(objRS("username"))
		objUtente.setPassword(objRS("pwd"))
		objUtente.setEmail(objRS("email"))
		objUtente.setPrivacy(objRS("privacy"))
		objUtente.setNewsletter(objRS("newsletter"))
		objUtente.setRuolo(objRS("ruolo"))
		objUtente.setUserActive(objRS("utenteAttivo"))
		objUtente.setSconto(objRS("sconto"))
		objUtente.setAdminComments(objRS("adminComments"))
		objUtente.setInsertDate(objRS("insertDate"))
		objUtente.setModifyDate(objRS("modifyDate"))
		objUtente.setPublic(objRS("public"))
		objUtente.setGroup(objRS("user_group"))
		objUtente.setAutomaticUser(objRS("automatic_user"))
		
		objDict.add strID, objUtente
		Set objUtente = nothing
		objRS.moveNext()
		loop
		
		Set getListaUtenti = objDict   
		Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function

	Public Function getListaUtentiNewsletter(id_newsletter)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		getListaUtentiNewsletter = ""  
		strSQL = "SELECT email FROM utenti INNER JOIN newsletter_x_utente ON utenti.id = newsletter_x_utente.id_utente WHERE newsletter = 1 AND email <> '' AND newsletter_x_utente.id_newsletter=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_newsletter)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then  
			Dim iCount, strEmailTmp
			strEmailTmp = ""
			
			do while not objRS.EOF
				strEmailTmp = strEmailTmp & objRS("email") & ";"
				objRS.moveNext()
			loop
			   
			getListaUtentiNewsletter = strEmailTmp   
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
	
	Public Function insertUser(strUserName, strPwd, strMail, strRuolo, bolPrivacy, bolNewsletter, bolUserActive, numSconto, strAdminComments, dateInsertDate, dateModifyDate, bolPublic, numUserGroup, bolAutomaticUser, objConn)
		on error resume next
		insertUser = -1
		
		Dim objDB, strSQL, strSQLSelect, objRS		

		if (Application("dbType") = 1) then
			dateInsertDate = convertDateTime(dateInsertDate)
			dateModifyDate = convertDateTime(dateModifyDate)
		end if

		strSQL = "INSERT INTO utenti(username, pwd, email, ruolo, privacy, newsletter, utenteAttivo, sconto, adminComments, insertDate,modifyDate,public, user_group, automatic_user) VALUES("
		strSQL = strSQL & "?,MD5(?),?,?,?,?,?,?,?,?,?,?,"
		if(isNull(numUserGroup) OR numUserGroup = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if		
		strSQL = strSQL & "?);"

		strSQLSelect = "SELECT username FROM utenti WHERE username =?"
		if not(strMail = "") then
			strSQLSelect = strSQLSelect & " AND email =?"
		end if
		strSQLSelect = strSQLSelect & ";"

		Dim objCommand, objCommand2
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQLSelect
		
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strUserName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strPwd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strMail)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strRuolo)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,bolPrivacy)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,bolNewsletter)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,bolUserActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numSconto))
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAdminComments)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dateInsertDate)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dateModifyDate)		
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolPublic)
		if not isNull(numUserGroup) AND not(numUserGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numUserGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolAutomaticUser)		
		
		objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,100,strUserName)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,100,strMail)
		Set objRS = objCommand2.Execute()
	
		if (objRS.EOF) then		
			objCommand.Execute()	
			Set objRS = objConn.Execute("SELECT max(utenti.id) as id FROM utenti;")
			if not (objRS.EOF) then
				insertUser = objRS("id")	
			end if			
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=001")
		end if
		
		Set objRS = Nothing	
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
			
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function

	Public Sub modifyUser(id, strUserName, strPwd, strMail, strRuolo, bolPrivacy, bolNewsletter, bolUserActive, numSconto, strAdminComments, dateInsertDate, dateModifyDate, bolPublic, numUserGroup, bolAutomaticUser, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		if (Application("dbType") = 1) then
			dateInsertDate = convertDateTime(dateInsertDate)
			dateModifyDate = convertDateTime(dateModifyDate)
		end if
		
		strSQL = "UPDATE utenti SET "
		strSQL = strSQL & "username=?,"
		if not(isNull(strPwd)) AND (strPwd <> "")then
			strSQL = strSQL & "pwd=MD5(?),"
		end if
		strSQL = strSQL & "email=?,"
		strSQL = strSQL & "privacy=?,"
		strSQL = strSQL & "newsletter=?,"
		strSQL = strSQL & "ruolo=?,"
		strSQL = strSQL & "utenteAttivo=?,"
		strSQL = strSQL & "sconto=?,"
		strSQL = strSQL & "adminComments=?,"
		strSQL = strSQL & "insertDate=?,"
		strSQL = strSQL & "modifyDate=?,"
		strSQL = strSQL & "public=?,"
		if(isNull(numUserGroup) OR numUserGroup = "") then
			strSQL = strSQL & "user_group=NULL,"
		else
			strSQL = strSQL & "user_group=?,"
		end if
		strSQL = strSQL & "automatic_user=?"
		strSQL = strSQL & " WHERE id=?;"		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strUserName)
		if not isNull(strPwd) AND not(strPwd = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strPwd)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strMail)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,bolPrivacy)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,bolNewsletter)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strRuolo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,bolUserActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numSconto))
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAdminComments)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dateInsertDate)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dateModifyDate)		
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolPublic)
		if not isNull(numUserGroup) AND not(numUserGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numUserGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolAutomaticUser)		
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)			
		
		Dim tmpObjUser, tmpStrMail
		Set tmpObjUser = findUserByID(id)
		tmpStrMail = tmpObjUser.getEmail()
		Set tmpObjUser = nothing		

		if (not(strMail = "") AND not(strMail = tmpStrMail)) then
			strSQLSelect = "SELECT username FROM utenti WHERE email =?;"
		
			Dim objCommand2
			Set objCommand2 = Server.CreateObject("ADODB.Command")
			objCommand2.ActiveConnection = objConn
			objCommand2.CommandType=1
			objCommand2.CommandText = strSQLSelect
			objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,100,strMail)
			Set objRS = objCommand2.Execute()
			
			if objRS.EOF then
				objCommand.Execute()
			else		
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=001&id_usr="&id)
			end if
		else
			objCommand.Execute()
		end if
		Set objCommand = Nothing
		Set objCommand2 = Nothing		
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
	
	Public Function insertUserNoTransaction(strUserName, strPwd, strMail, strRuolo, bolPrivacy, bolNewsletter, bolUserActive, numSconto, strAdminComments, dateInsertDate, dateModifyDate, bolPublic, numUserGroup, bolAutomaticUser)
		on error resume next
		insertUserNoTransaction = -1
		
		Dim objDB, strSQL, strSQLSelect, objRS, objConn		

		if (Application("dbType") = 1) then
			dateInsertDate = convertDateTime(dateInsertDate)
			dateModifyDate = convertDateTime(dateModifyDate)
		end if	

		strSQL = "INSERT INTO utenti(username, pwd, email, ruolo, privacy, newsletter, utenteAttivo, sconto, adminComments, insertDate,modifyDate,public, user_group, automatic_user) VALUES("
		strSQL = strSQL & "?,MD5(?),?,?,?,?,?,?,?,?,?,?,"
		if(isNull(numUserGroup) OR numUserGroup = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if		
		strSQL = strSQL & "?);"

		strSQLSelect = "SELECT username FROM utenti WHERE username =?"
		if not(strMail = "") then
			strSQLSelect = strSQLSelect & " AND email =?"
		end if
		strSQLSelect = strSQLSelect & ";"
		
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
		objCommand2.CommandText = strSQLSelect
		
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strUserName)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strPwd)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strMail)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strRuolo)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,bolPrivacy)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,bolNewsletter)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,bolUserActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numSconto))
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAdminComments)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dateInsertDate)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dateModifyDate)		
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolPublic)
		if not isNull(numUserGroup) AND not(numUserGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numUserGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolAutomaticUser)
		
		
		objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,100,strUserName)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,100,strMail)
		Set objRS = objCommand2.Execute()
		
		if (objRS.EOF) then	
			objCommand.Execute()		
			Set objRS = objConn.Execute("SELECT max(utenti.id) as id FROM utenti;")
			if not (objRS.EOF) then
				insertUserNoTransaction = objRS("id")	
			end if			
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=001")
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objDB = Nothing
			
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function

	Public Sub modifyUserNoTransaction(id, strUserName, strPwd, strMail, strRuolo, bolPrivacy, bolNewsletter, bolUserActive, numSconto, strAdminComments, dateInsertDate, dateModifyDate, bolPublic, numUserGroup, bolAutomaticUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		if (Application("dbType") = 1) then
			dateInsertDate = convertDateTime(dateInsertDate)
			dateModifyDate = convertDateTime(dateModifyDate)
		end if

		strSQL = "UPDATE utenti SET "
		strSQL = strSQL & "username=?,"
		if not(isNull(strPwd)) AND (strPwd <> "")then
			strSQL = strSQL & "pwd=MD5(?),"
		end if
		strSQL = strSQL & "email=?,"
		strSQL = strSQL & "privacy=?,"
		strSQL = strSQL & "newsletter=?,"
		strSQL = strSQL & "ruolo=?,"
		strSQL = strSQL & "utenteAttivo=?,"
		strSQL = strSQL & "sconto=?,"
		strSQL = strSQL & "adminComments=?,"
		strSQL = strSQL & "insertDate=?,"
		strSQL = strSQL & "modifyDate=?,"
		strSQL = strSQL & "public=?,"
		if(isNull(numUserGroup) OR numUserGroup = "") then
			strSQL = strSQL & "user_group=NULL,"
		else
			strSQL = strSQL & "user_group=?,"
		end if
		strSQL = strSQL & "automatic_user=?"
		strSQL = strSQL & " WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strUserName)
		if not isNull(strPwd) AND not(strPwd = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strPwd)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strMail)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,bolPrivacy)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,bolNewsletter)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strRuolo)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,bolUserActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numSconto))
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strAdminComments)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dateInsertDate)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dateModifyDate)		
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolPublic)
		if not isNull(numUserGroup) AND not(numUserGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,numUserGroup)
		end if	
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolAutomaticUser)	
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		
		
		Dim tmpObjUser, tmpStrMail
		Set tmpObjUser = findUserByID(id)
		tmpStrMail = tmpObjUser.getEmail()
		Set tmpObjUser = nothing		

		if (not(strMail = "") AND not(strMail = tmpStrMail)) then
			strSQLSelect = "SELECT username FROM utenti WHERE email =?;"
		
			Dim objCommand2
			Set objCommand2 = Server.CreateObject("ADODB.Command")
			objCommand2.ActiveConnection = objConn
			objCommand2.CommandType=1
			objCommand2.CommandText = strSQLSelect
			objCommand2.Parameters.Append objCommand2.CreateParameter(,200,1,100,strMail)
			Set objRS = objCommand2.Execute()
			
			if objRS.EOF then
				objCommand.Execute()
			else		
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=001&id_usr="&id)
			end if
		else
			objCommand.Execute()
		end if
		
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
		
	Public Function deleteUser(id)
		on error resume next		
		deleteUser = true
		
		Dim objDB, strSQL, objRS, objRS2, objConn, strSQL2, strSQL2B, strSQL3, strSQL4, strSQL5, strSQL6
		strSQL = "DELETE FROM utenti WHERE id=?;" 
		strSQL2 = "SELECT news_x_utente.id_utente FROM news_x_utente WHERE news_x_utente.id_utente=?;"
		strSQL2B = "SELECT ordini.id_utente FROM ordini WHERE ordini.id_utente=?;"
		strSQL3 = "DELETE FROM target_x_utente WHERE id_utente=?;"
		strSQL4 = "DELETE FROM newsletter_x_utente WHERE id_utente=?;"
		strSQL5 = "DELETE FROM utenti_images WHERE id_utente=?;"
		strSQL6 = "DELETE FROM friend_x_utente WHERE id_user=?;"
		strSQL7 = "DELETE FROM user_fields_match WHERE id_user=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	

		Dim objCommand, objCommand2, objCommand3, objCommand4, objCommand5, objCommand6, objCommand7, objCommand8
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		Set objCommand3 = Server.CreateObject("ADODB.Command")
		Set objCommand4 = Server.CreateObject("ADODB.Command")
		Set objCommand5 = Server.CreateObject("ADODB.Command")
		Set objCommand6 = Server.CreateObject("ADODB.Command")
		Set objCommand7 = Server.CreateObject("ADODB.Command")
		Set objCommand8 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand3.ActiveConnection = objConn
		objCommand4.ActiveConnection = objConn
		objCommand5.ActiveConnection = objConn
		objCommand6.ActiveConnection = objConn
		objCommand7.ActiveConnection = objConn
		objCommand8.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand3.CommandType=1
		objCommand4.CommandType=1
		objCommand5.CommandType=1
		objCommand6.CommandType=1
		objCommand7.CommandType=1
		objCommand8.CommandType=1
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQL2
		objCommand3.CommandText = strSQL2B
		objCommand4.CommandText = strSQL3
		objCommand5.CommandText = strSQL4
		objCommand6.CommandText = strSQL5
		objCommand7.CommandText = strSQL6
		objCommand8.CommandText = strSQL7
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)
		objCommand4.Parameters.Append objCommand4.CreateParameter(,19,1,,id)
		objCommand5.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand6.Parameters.Append objCommand6.CreateParameter(,19,1,,id)
		objCommand7.Parameters.Append objCommand7.CreateParameter(,19,1,,id)
		objCommand8.Parameters.Append objCommand8.CreateParameter(,19,1,,id)

		objConn.BeginTrans
		
		Set objRS = objCommand2.Execute()
'<!--nsys-objusr1-->
		Set objRS2 = objCommand3.Execute()
		if not(objRS.EOF) OR not(objRS2.EOF) then							
'<!---nsys-objusr1-->
			deleteUser = false				
		else
			if(Application("use_innodb_table") = 0) then
				objCommand4.Execute()
				objCommand5.Execute()
				objCommand6.Execute()
				objCommand7.Execute()
				objCommand8.Execute()
			end if	
			objCommand.Execute()			
		end if
		
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		Set objCommand4 = Nothing
		Set objCommand5 = Nothing
		Set objCommand6 = Nothing
		Set objCommand7 = Nothing
		Set objCommand8 = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function

	Public Sub disableUser(id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE utenti SET "
		strSQL = strSQL & "utenteAttivo=0"
		strSQL = strSQL & " WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Execute()
		Set objCommand = Nothing	
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub changePassword(id, strPwd)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE utenti SET "
		strSQL = strSQL & "pwd=MD5(?)"
		strSQL = strSQL & " WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strPwd)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing	
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Function findUserByID(id)
		on error resume next
		
		findUserByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM utenti WHERE id=?;"

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
			Dim objUser
			Set objUser = new UserClass
			objUser.setUserID(objRS("id"))
			objUser.setUserName(objRS("username"))
			objUser.setPassword(objRS("pwd"))
			objUser.setEmail(objRS("email"))
			objUser.setRuolo(objRS("ruolo"))
			objUser.setPrivacy(objRS("privacy"))
			objUser.setNewsletter(objRS("newsletter"))		
			objUser.setUserActive(objRS("utenteAttivo"))
			objUser.setSconto(objRS("sconto"))
			objUser.setAdminComments(objRS("adminComments"))
			objUser.setInsertDate(objRS("insertDate"))
			objUser.setModifyDate(objRS("modifyDate"))
			objUser.setPublic(objRS("public"))
			objUser.setGroup(objRS("user_group"))
			objUser.setAutomaticUser(objRS("automatic_user"))

			Set findUserByID = objUser
			Set objUser = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if				
	End Function
		
	Public Function findUserByIDExt(id, redirectOnError)
		on error resume next
		
		findUserByIDExt = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM utenti WHERE id=?;"

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
			if(redirectOnError) then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")	
			end if	
		else
			Dim objUser
			Set objUser = new UserClass
			objUser.setUserID(objRS("id"))
			objUser.setUserName(objRS("username"))
			objUser.setPassword(objRS("pwd"))
			objUser.setEmail(objRS("email"))
			objUser.setRuolo(objRS("ruolo"))
			objUser.setPrivacy(objRS("privacy"))
			objUser.setNewsletter(objRS("newsletter"))		
			objUser.setUserActive(objRS("utenteAttivo"))
			objUser.setSconto(objRS("sconto"))
			objUser.setAdminComments(objRS("adminComments"))
			objUser.setInsertDate(objRS("insertDate"))
			objUser.setModifyDate(objRS("modifyDate"))
			objUser.setPublic(objRS("public"))
			objUser.setGroup(objRS("user_group"))
			objUser.setAutomaticUser(objRS("automatic_user"))

			Set findUserByIDExt = objUser
			Set objUser = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if			
	End Function
		
	Public Function findUserByUserID(strUser)
		on error resume next
		
		findUserByUserID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM utenti WHERE username=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strUser)
		Set objRS = objCommand.Execute()		

		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")		
		else
			Dim objUser
			Set objUser = new UserClass
			objUser.setUserID(objRS("id"))
			objUser.setUserName(objRS("username"))
			objUser.setPassword(objRS("pwd"))
			objUser.setEmail(objRS("email"))
			objUser.setRuolo(objRS("ruolo"))
			objUser.setPrivacy(objRS("privacy"))
			objUser.setNewsletter(objRS("newsletter"))		
			objUser.setUserActive(objRS("utenteAttivo"))
			objUser.setSconto(objRS("sconto"))
			objUser.setAdminComments(objRS("adminComments"))
			objUser.setInsertDate(objRS("insertDate"))
			objUser.setModifyDate(objRS("modifyDate"))
			objUser.setPublic(objRS("public"))
			objUser.setGroup(objRS("user_group"))
			objUser.setAutomaticUser(objRS("automatic_user"))
			
			Set findUserByUserID = objUser
			Set objUser = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findUserByUserAndMail(strUser, strMail)
		on error resume next
		
		findUserByUserAndMail = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM utenti WHERE username=? AND email=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strMail)
		Set objRS = objCommand.Execute()		

		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")		
		else
			Dim objUser
			Set objUser = new UserClass
			objUser.setUserID(objRS("id"))
			objUser.setUserName(objRS("username"))
			objUser.setPassword(objRS("pwd"))
			objUser.setEmail(objRS("email"))
			objUser.setRuolo(objRS("ruolo"))
			objUser.setPrivacy(objRS("privacy"))
			objUser.setNewsletter(objRS("newsletter"))			
			objUser.setUserActive(objRS("utenteAttivo"))
			objUser.setSconto(objRS("sconto"))
			objUser.setAdminComments(objRS("adminComments"))
			objUser.setInsertDate(objRS("insertDate"))
			objUser.setModifyDate(objRS("modifyDate"))
			objUser.setPublic(objRS("public"))
			objUser.setGroup(objRS("user_group"))
			objUser.setAutomaticUser(objRS("automatic_user"))
			
			Set findUserByUserAndMail = objUser
			Set objUser = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	

	Public Function login(strUser, strPwd, paramFrom)
		on error resume next
		
		login = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM utenti WHERE username=? AND pwd=MD5(?) AND utenteAttivo='1';"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strPwd)
		Set objRS = objCommand.Execute()		

		if objRS.EOF then
			response.Redirect(Application("baseroot")&"/login.asp?error=003&from="&paramFrom)		
		else
			Dim objUser
			Set objUser = new UserClass
			objUser.setUserID(objRS("id"))
			objUser.setUserName(objRS("username"))
			objUser.setPassword(objRS("pwd"))
			objUser.setEmail(objRS("email"))
			objUser.setRuolo(objRS("ruolo"))
			objUser.setPrivacy(objRS("privacy"))
			objUser.setNewsletter(objRS("newsletter"))		
			objUser.setUserActive(objRS("utenteAttivo"))
			objUser.setSconto(objRS("sconto"))
			objUser.setAdminComments(objRS("adminComments"))
			objUser.setInsertDate(objRS("insertDate"))
			objUser.setModifyDate(objRS("modifyDate"))
			objUser.setPublic(objRS("public"))
			objUser.setGroup(objRS("user_group"))
			objUser.setAutomaticUser(objRS("automatic_user"))
						
			Set login = objUser
			Set objUser = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Sub insertUserXNews(id_user, id_news, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO news_x_utente(id_news, id_utente) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
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

	Public Sub insertUserXNewsNoTransaction(id_user, id_news)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO news_x_utente(id_news, id_utente) VALUES("
		strSQL = strSQL & "?,?);"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_news)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub insertTargetXUser(id_target, id_user, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO target_x_utente(id_target, id_utente) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
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
	
	Public Sub insertTargetXUserNoTransaction(id_target, id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO target_x_utente(id_target, id_utente) VALUES("
		strSQL = strSQL & "?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteTargetXUser(id_user, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM target_x_utente WHERE id_utente=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
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
		
	Public Sub deleteTargetXUserNoTransaction(id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM target_x_utente WHERE id_utente=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub


	'************************* GESTIONE LISTA FRIENDS *******************************
		
	Public Function getListaFriends(idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objUtente
		getListaFriends = null  
		strSQL = "SELECT utenti.*, friend_x_user.active FROM utenti LEFT JOIN friend_x_user ON (friend_x_user.id_friend=utenti.id) WHERE id_user=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		Set objRS = objCommand.Execute()	

		if not(objRS.EOF) then   
		Set objDict = Server.CreateObject("Scripting.Dictionary")   
		do while not objRS.EOF
		Set objUtente = new UserClass
		strID = objRS("id")
		
		objUtente.setUserID(strID)    
		objUtente.setUserName(objRS("username"))
		objUtente.setPassword(objRS("pwd"))
		objUtente.setEmail(objRS("email"))
		objUtente.setPrivacy(objRS("privacy"))
		objUtente.setNewsletter(objRS("newsletter"))
		objUtente.setRuolo(objRS("ruolo"))
		objUtente.setUserActive(objRS("utenteAttivo"))
		objUtente.setFriendActive(objRS("active"))
		objUtente.setSconto(objRS("sconto"))
		objUtente.setAdminComments(objRS("adminComments"))
		objUtente.setInsertDate(objRS("insertDate"))
		objUtente.setModifyDate(objRS("modifyDate"))
		objUtente.setPublic(objRS("public"))
		objUtente.setGroup(objRS("user_group"))
		objUtente.setAutomaticUser(objRS("automatic_user"))
		
		objDict.add strID, objUtente
		objRS.moveNext()
		loop
		
		Set getListaFriends = objDict   
		Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function

	Public Function bolHasFriend(idFriend, idUser)
		on error resume next
		
		bolHasFriend = false
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM friend_x_user WHERE id_friend=? AND id_user=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			bolHasFriend = false	
		else
			bolHasFriend = true
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if			
	End Function

	Public Function bolHasFriendActive(idFriend, idUser)
		on error resume next
		
		bolHasFriendActive = false
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM friend_x_user WHERE id_friend=? AND id_user=? AND active=1;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			bolHasFriendActive = false	
		else
			bolHasFriendActive = true
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			bolHasFriendActive = false
		end if			
	End Function
	
	Public Sub insertFriendXUser(idFriend, idUser, isActive, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO friend_x_user(id_friend, id_user, active) VALUES("
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,isActive)
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
	
	Public Sub insertFriendXUserNoTransaction(idFriend, idUser, isActive)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO friend_x_user(id_friend, id_user, active) VALUES("
		strSQL = strSQL & "?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,isActive)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub updateFriendStatus(idFriend, idUser, isActive, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "UPDATE friend_x_user SET active=? WHERE id_friend=? AND id_user=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,isActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
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
	
	Public Sub updateFriendStatusNoTransaction(idFriend, idUser, isActive)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE friend_x_user SET active=? WHERE id_friend=? AND id_user=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,isActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing	
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteFriendXUser(idUser, idFriend, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM friend_x_user WHERE id_user=? AND id_friend=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
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
		
	Public Sub deleteFriendXUserNoTransaction(idUser, idFriend)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM friend_x_user WHERE id_user=? AND id_friend=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	'************************* FINE GESTIONE LISTA FRIENDS *******************************
		
	Public Function getListaRuoli()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaRuoli = null		
		strSQL = "SELECT * FROM ruoli_utente;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("descrizione")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaRuoli = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getMaxIDUtente()
		on error resume next
		
		getMaxIDUtente = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT max(utenti.id) as id FROM utenti;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDUtente = objRS("id")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getTargetPerUser(id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getTargetPerUser = null		
		strSQL = "SELECT target_x_utente.id_target, target.descrizione, target.type, target.automatic FROM target INNER JOIN target_x_utente ON target.id = target_x_utente.id_target WHERE target_x_utente.id_utente=? ORDER BY target.type, target.descrizione;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		Set objRS = objCommand.Execute()	

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objTarget = new Targetclass		
				strID = objRS("id_target")
				objTarget.setTargetID(strID)
				objTarget.setTargetDescrizione(objRS("descrizione"))
				objTarget.setTargetType(objRS("type"))	
				objTarget.setAutomatic(objRS("automatic"))
				objDict.add strID, objTarget
				Set objTarget = nothing
				objRS.moveNext()
			loop
							
			Set getTargetPerUser = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function		
	
	Public Sub insertNewsletterXUser(id_newsletter, id_user, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO newsletter_x_utente(id_newsletter, id_utente) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_newsletter)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
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
		
	Public Sub deleteNewsletterXUser(id_user, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM newsletter_x_utente WHERE id_utente=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
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
	
	Public Sub insertNewsletterXUserNoTransaction(id_newsletter, id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO newsletter_x_utente(id_newsletter, id_utente) VALUES("
		strSQL = strSQL & "?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_newsletter)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteNewsletterXUserNoTransaction(id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM newsletter_x_utente WHERE id_utente=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Function getNewsletterPerUser(id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getNewsletterPerUser = null		
		strSQL = "SELECT * FROM newsletter_x_utente WHERE id_utente=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id_newsletter")
				strIDUsr = objRS("id_utente")		
				objDict.add strID, strIDUsr
				objRS.moveNext()
			loop
							
			Set getNewsletterPerUser = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	
	'******************************* FUNZIONI PER CONFERMA UTENTE *******************************
	Public Function findConfirmationCodeUserByID(id_user, confirmation_code)
		on error resume next
		
		findConfirmationCodeUserByID = false
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM conferma_utente WHERE id_user=? AND confirmation_code=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,confirmation_code)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			findConfirmationCodeUserByID = false		
		else
			findConfirmationCodeUserByID = true
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertConfirmationCodeXUser(id_user, confirmation_code, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO conferma_utente(id_user, confirmation_code) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,confirmation_code)
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
	
	Public Sub insertConfirmationCodeXUserNoTransaction(id_user, confirmation_code)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO conferma_utente(id_user, confirmation_code) VALUES("
		strSQL = strSQL & "?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_user)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,confirmation_code)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub activateUser(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE utenti SET "
		strSQL = strSQL & "utenteAttivo='1'"
		strSQL = strSQL & " WHERE id=?;"
		
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
	
	'******************************* FUNZIONI PER IMAGE UTENTE *******************************
	
	Public Function findImageIDXUser(id_user)
		on error resume next
		
		findImageIDXUser = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT id FROM utenti_images WHERE id_utente=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then
			findImageIDXUser = objRS("id")		
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function hasImageUser(id_user)
		on error resume next
		
		hasImageUser = false
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT id FROM utenti_images WHERE id_utente=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then
			hasImageUser = true	
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertImageXUser(id_utente, filename, content_type, file_size, file_data, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		' Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open "utenti_images",objConn, 2, 2, 2

		' Adding data
		objRS.AddNew
		objRS("id_utente") = id_utente
		objRS("filename") = filename
		objRS("content_type") = content_type
		objRS("file_size") = file_size
		objRS("file_data").AppendChunk file_data
		objRS.Update
		objRS.Close
		Set objRS = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub updateImageXUser(id, id_utente, filename, content_type, file_size, file_data, objConn)
		on error resume next
		Dim objDB, objRS
				
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open "utenti_images",objConn, 1, 2, adCmdTableDirect
		objRS.Filter = "id="&id
		objRS("id_utente") = id_utente
		objRS("filename") = filename
		objRS("content_type") = content_type
		objRS("file_size") = file_size
		objRS("file_data").AppendChunk file_data
		objRS.Update
		objRS.Close
		Set objRS = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub deleteImageXUser(id_utente, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "DELETE FROM utenti_images WHERE id_utente=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_utente)
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
	
	Public Sub insertImageXUserNoTransaction(id_utente, filename, content_type, file_size, file_data)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()

		' Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open "utenti_images",objConn, 2, 2, 2

		' Adding data
		objRS.AddNew
		objRS("id_utente") = id_utente
		objRS("filename") = filename
		objRS("content_type") = content_type
		objRS("file_size") = file_size
		objRS("file_data").AppendChunk file_data
		objRS.Update
		objRS.Close
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub updateImageXUserNoTransaction(id, id_utente, filename, content_type, file_size, file_data)
		on error resume next
		Dim objDB, objRS, objConn
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open "utenti_images",objConn, 1, 2, adCmdTableDirect	
		objRS.Filter = "id="&id
		objRS("id_utente") = id_utente
		objRS("filename") = filename
		objRS("content_type") = content_type
		objRS("file_size") = file_size
		objRS("file_data").AppendChunk file_data
		objRS.Update
		objRS.Close
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Sub deleteImageXUserNoTransaction(id_utente)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "DELETE FROM utenti_images WHERE id_utente=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_utente)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Function getUserImageData(id_utente)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		getUserImageData = null
		
		strSQL = "SELECT file_data FROM utenti_images WHERE id_utente=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_utente)
		Set objRS = objCommand.Execute()		

		'GET binary data from recordset
		getUserImageData = objRS("file_data")

		'Use this code instead of previous line For ORACLE.
		'getUserImageData = objRS("file_data").GetChunk(objRS("file_data").ActualSize)
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
	
	Function getUserImageObjectNoData(id_utente)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		Dim idFile, idUtente, strFilename, strFileSize, strFileContentType
		
		getUserImageObjectNoData = null
		
		strSQL = "SELECT id,id_utente,filename,content_type,file_size FROM utenti_images WHERE id_utente=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_utente)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			idFile = objRS("id")
			idUtente = objRS("id_utente")
			strFilename = objRS("filename")
			strFileSize =objRS("file_size")
			strFileContentType = objRS("content_type")
			
			objDict.add "id", idFile
			objDict.add "id_utente", idUtente
			objDict.add "filename", strFilename
			objDict.add "content_type", strFileContentType
			objDict.add "file_size", strFileSize
			'objDict.add "file_data", objRS("file_data")	
			'Use this code instead of previous line For ORACLE.
			'objDict.add "file_data", objRS("file_data").GetChunk(objRS("file_data").ActualSize)
		
			Set getUserImageObjectNoData = objDict			
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
		
		convertDate = YY&"-"&MM&"-"&DD		
	End Function
	
	Public Function convertDateTime(dateToConvert)
		Dim DD, MM, YY, HH, MIN, SS
		
		convertDateTime = null
		
		DD = DatePart("d", dateToConvert)
		MM = DatePart("m", dateToConvert)
		YY = DatePart("yyyy", dateToConvert)
		HH = DatePart("h", dateToConvert)
		MIN = DatePart("n", dateToConvert)
		SS = DatePart("s", dateToConvert)
		
		convertDateTime = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS		
	End Function

	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = doubleValue
		
		'if (Application("dbType") = 0) then
			convertDoubleDelimiter = Replace(convertDoubleDelimiter, ".",",")
		'else		
			'convertDoubleDelimiter = Replace(convertDoubleDelimiter, ",",".")
		'end if			
	End Function

	Function SortDictionary(objDict,intSort)
	  ' declare our variables
	  Dim dictKey, dictItem
	  Dim strDict()
	  Dim objKey
	  Dim strKey,strItem
	  Dim X,Y,Z
	  
	  'Set SortDictionary = null
	  
	  dictKey  = 1
	  dictItem = 2
	
	  ' get the dictionary count
	  Z = objDict.Count
	
	  ' we need more than one item to warrant sorting
	  If Z > 1 Then
		' create an array to store dictionary information
		ReDim strDict(Z,2)
		X = 0
		' populate the string array
		For Each objKey In objDict
			strDict(X,dictKey)  = CStr(objKey)
			strDict(X,dictItem) = CStr(objDict(objKey))
			X = X + 1
		Next
	
		' perform a a shell sort of the string array
		For X = 0 to (Z - 2)
		  For Y = X to (Z - 1)
			If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
				strKey  = strDict(X,dictKey)
				strItem = strDict(X,dictItem)
				strDict(X,dictKey)  = strDict(Y,dictKey)
				strDict(X,dictItem) = strDict(Y,dictItem)
				strDict(Y,dictKey)  = strKey
				strDict(Y,dictItem) = strItem
			End If
		  Next
		Next
	
		' erase the contents of the dictionary object
		objDict.RemoveAll
	
		' repopulate the dictionary with the sorted information
		For X = 0 to (Z - 1)
		  objDict.Add strDict(X,dictKey), strDict(X,dictItem)
		Next
	
	  End If
	  Set SortDictionary = objDict
	End Function
	
	public Sub toString()
		response.write (idUser & ", " & usrName & ", " & pwd & ", " & email & ", " & ruolo & ", " & privacy & ", " & newsletter)
	end Sub
End Class
%>