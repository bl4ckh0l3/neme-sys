<%
Class UserPreferenceClass
	Private id
	Private idUser
	Private idFriend
	Private idCommentoUser
	Private typeCommento
	Private tipo
	Private value
	Private insertDate
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strid)
		id = strid
	End Sub	
	
	Public Function getIdUser()
		getIdUser = idUser
	End Function
	
	Public Sub setIdUser(stridUser)
		idUser = stridUser
	End Sub
	
	Public Function getIdFriend()
		getIdFriend = idFriend
	End Function
	
	Public Sub setIdFriend(stridFriend)
		idFriend = stridFriend
	End Sub	
	
	Public Function getIdCommentoUser()
		getIdCommentoUser = idCommentoUser
	End Function
	
	Public Sub setIdCommentoUser(stridCommentoUser)
		idCommentoUser = stridCommentoUser
	End Sub
	
	Public Function getTypeCommento()
		getTypeCommento = typeCommento
	End Function
	
	Public Sub setTypeCommento(strTypeCommento)
		typeCommento = strTypeCommento
	End Sub
	
	Public Function getType()
		getType = tipo
	End Function
	
	Public Sub setType(strType)
		tipo = strType
	End Sub
	
	Public Function getValue()
		getValue = value
	End Function
	
	Public Sub setValue(strValue)
		value = strValue
	End Sub
	
	Public Function getInsertDate()
		getInsertDate = insertDate
	End Function
	
	Public Sub setInsertDate(dateInsertDate)
		insertDate = dateInsertDate
	End Sub
		
	Public Function getListUserPreferenceByUserFiltered(idUser,excludeComment, excludeCommentType)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListUserPreferenceByUserFiltered = null		
		strSQL = "SELECT * FROM user_preference WHERE id_user=?"
		
		if(excludeComment) then
			strSQL = strSQL& " AND (ISNULL(id_usr_comment) OR id_usr_comment='')"
		else
			strSQL = strSQL& " AND (NOT ISNULL(id_usr_comment) AND id_usr_comment<>'')"
		end if		
		
		if(excludeCommentType) then
			strSQL = strSQL& " AND (ISNULL(comment_type) OR comment_type<>'')"
		else
			strSQL = strSQL& " AND (NOT ISNULL(comment_type) AND comment_type<>'')"
		end if
		
		strSQL = strSQL & " ORDER BY dta_insert DESC, id_usr_comment DESC;"	
		
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
			Dim objUserPreference
			do while not objRS.EOF				
				Set objUserPreference = new UserPreferenceClass
				strID = objRS("id")
				objUserPreference.setID(strID)
				objUserPreference.setIdUser(objRS("id_user"))
				objUserPreference.setIdFriend(objRS("id_friend"))
				objUserPreference.setType(objRS("type"))	
				objUserPreference.setValue(objRS("value"))	
				objUserPreference.setInsertDate(objRS("dta_insert"))	
				objUserPreference.setIdCommentoUser(objRS("id_usr_comment"))	
				objUserPreference.setTypeCommento(objRS("comment_type"))				
				objDict.add strID, objUserPreference
				objRS.moveNext()
			loop
			Set objUserPreference = nothing							
			Set getListUserPreferenceByUserFiltered = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListUserPreferenceByFriend(idUser, idFriend)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListUserPreferenceByFriend = null		
		strSQL = "SELECT * FROM user_preference WHERE id_user=? AND id_friend=? ORDER BY dta_insert DESC, id_usr_comment DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		Set objRS = objCommand.Execute()				

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objUserPreference
			do while not objRS.EOF				
				Set objUserPreference = new UserPreferenceClass
				strID = objRS("id")
				objUserPreference.setID(strID)
				objUserPreference.setIdUser(objRS("id_user"))
				objUserPreference.setIdFriend(objRS("id_friend"))
				objUserPreference.setType(objRS("type"))	
				objUserPreference.setValue(objRS("value"))	
				objUserPreference.setInsertDate(objRS("dta_insert"))	
				objUserPreference.setIdCommentoUser(objRS("id_usr_comment"))	
				objUserPreference.setTypeCommento(objRS("comment_type"))	
				objDict.add strID, objUserPreference
				objRS.moveNext()
			loop
			Set objUserPreference = nothing							
			Set getListUserPreferenceByFriend = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

		
	Public Function getListUserPreferenceByFriendAndComment(idUser, idFriend, idComment)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListUserPreferenceByFriendAndComment = null		
		strSQL = "SELECT * FROM user_preference WHERE id_user=?"
		if not(isNull(idFriend)) then strSQL = strSQL & " AND id_friend=?"
		strSQL = strSQL &" AND id_usr_comment=? ORDER BY dta_insert DESC, id_usr_comment DESC;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		if not(isNull(idFriend)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)		
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idComment)
		Set objRS = objCommand.Execute()				

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objUserPreference
			do while not objRS.EOF				
				Set objUserPreference = new UserPreferenceClass
				strID = objRS("id")
				objUserPreference.setID(strID)
				objUserPreference.setIdUser(objRS("id_user"))
				objUserPreference.setIdFriend(objRS("id_friend"))
				objUserPreference.setType(objRS("type"))	
				objUserPreference.setValue(objRS("value"))	
				objUserPreference.setInsertDate(objRS("dta_insert"))	
				objUserPreference.setIdCommentoUser(objRS("id_usr_comment"))	
				objUserPreference.setTypeCommento(objRS("comment_type"))	
				objDict.add strID, objUserPreference
				Set objUserPreference = nothing				
				objRS.moveNext()
			loop			
			Set getListUserPreferenceByFriendAndComment = objDict		
			Set objDict = nothing				
		end if
	
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing

		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findNumUserPreferenceByType(tipo, idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findNumUserPreferenceByType = null		
		strSQL = "SELECT count(`type`) AS counter FROM user_preference WHERE `type`=? AND id_user=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		Set objRS = objCommand.Execute()				

		if not(objRS.EOF) then			
			findNumUserPreferenceByType = objRS("counter")		
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function findNumUserPreferenceTotal(idUser, exclude)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findNumUserPreferenceTotal = null		
		strSQL = "SELECT count(type) AS counter FROM user_preference WHERE id_user=?"

		if(exclude)then
		strSQL = strSQL & " AND type <> -1"
		end if
		strSQL = strSQL & ";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idUser)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			findNumUserPreferenceTotal = objRS("counter")		
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function		

	Public Function findUserPreferencePositivePercent(idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findUserPreferencePositivePercent = 0		

		Dim positive, total
		positive = findNumUserPreferenceByType(1, idUser)
		total = findNumUserPreferenceTotal(idUser, true)

		if(Cint(total) > 0) then
			findUserPreferencePositivePercent = (Cint(positive) * 100) / Cint(total)
		end if
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Sub insertUserPreference(idUser, idFriend, idUsrComment, commentType, tipo, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, dtaInsert

		dtaInsert = now()
		if (Application("dbType") = 1) then
			dtaInsert = convertDateTime(dtaInsert)
		end if
	
		strSQL = "INSERT INTO user_preference(id_user, id_friend, id_usr_comment, comment_type, `type`, value, dta_insert) VALUES("
		strSQL = strSQL & "?,?,"

		if(isNull(idUsrComment) OR idUsrComment = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if

		if(isNull(commentType) OR commentType = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if	
		
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		if not isNull(idUsrComment) AND not(idUsrComment = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUsrComment)
		end if
		if not isNull(commentType) AND not(commentType = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,commentType)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,value)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtaInsert)
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
		
	Public Sub modifyUserPreference(id, idUser, idFriend, idUsrComment, commentType, tipo, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS, dtaModify
		
		dtaModify = now()
		if (Application("dbType") = 1) then
			dtaModify = convertDateTime(dtaModify)
		end if

		strSQL = "UPDATE user_preference SET "
		strSQL = strSQL & "id_user=?,"
		strSQL = strSQL & "id_friend=?,"
		if(isNull(idUsrComment) OR idUsrComment = "") then
			strSQL = strSQL & "id_usr_comment=NULL,"
		else
			strSQL = strSQL & "id_usr_comment=?,"			
		end if
		if(isNull(commentType) OR commentType = "") then
			strSQL = strSQL & "comment_type=NULL,"
		else
			strSQL = strSQL & "comment_type=?,"			
		end if
		strSQL = strSQL & "`type`=?,"		
		strSQL = strSQL & "value=?,"
		strSQL = strSQL & "dta_insert=?"
		strSQL = strSQL & " WHERE id=?;" 

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		if not isNull(idUsrComment) AND not(idUsrComment = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUsrComment)
		end if
		if not isNull(commentType) AND not(commentType = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,commentType)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,value)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtaModify)
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
			
	Public Sub insertUserPreferenceNoTransaction(idUser, idFriend, idUsrComment, commentType, tipo, value)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, dtaInsert
		
		dtaInsert = now()
		if (Application("dbType") = 1) then
			dtaInsert = convertDateTime(dtaInsert)
		end if
		
		strSQL = "INSERT INTO user_preference(id_user, id_friend, id_usr_comment, comment_type, `type`, value, dta_insert) VALUES("
		strSQL = strSQL & "?,?,"

		if(isNull(idUsrComment) OR idUsrComment = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if

		if(isNull(commentType) OR commentType = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if	
		
		strSQL = strSQL & "?,?,?);"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		if not isNull(idUsrComment) AND not(idUsrComment = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUsrComment)
		end if
		if not isNull(commentType) AND not(commentType = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,commentType)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,value)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtaInsert)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyUserPreferenceNoTransaction(id, idUser, idFriend, idUsrComment, commentType, tipo, value)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, dtaModify
		
		dtaModify = now()
		if (Application("dbType") = 1) then
			dtaModify = convertDateTime(dtaModify)
		end if

		strSQL = "UPDATE user_preference SET "
		strSQL = strSQL & "id_user=?,"
		strSQL = strSQL & "id_friend=?,"
		if(isNull(idUsrComment) OR idUsrComment = "") then
			strSQL = strSQL & "id_usr_comment=NULL,"
		else
			strSQL = strSQL & "id_usr_comment=?,"			
		end if
		if(isNull(commentType) OR commentType = "") then
			strSQL = strSQL & "comment_type=NULL,"
		else
			strSQL = strSQL & "comment_type=?,"			
		end if
		strSQL = strSQL & "`type`=?,"		
		strSQL = strSQL & "value=?,"
		strSQL = strSQL & "dta_insert=?"
		strSQL = strSQL & " WHERE id=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idFriend)
		if not isNull(idUsrComment) AND not(idUsrComment = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUsrComment)
		end if
		if not isNull(commentType) AND not(commentType = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,commentType)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,value)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dtaModify)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteUserPreference(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM user_preference WHERE id=?;" 

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
		
	Public Sub deleteUserPreferenceByUser(idUser)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM user_preference WHERE id_user=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idUser)
		objCommand.Execute()	
		Set objCommand = Nothing	
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteUserPreferenceByFriend(idUser, idFriend)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM user_preference WHERE id_user=? AND id_friend=?;"

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
End Class
%>