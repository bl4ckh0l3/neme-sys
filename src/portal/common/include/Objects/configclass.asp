<%

Class ConfigClass
	Private key
	Private descrizione
	Private str_value
	Private alert
	Private str_type
	
	
	Public Function getKey()
		getKey = key
	End Function
	
	Public Sub setKey(strKey)
		key = strKey
	End Sub
	
	Public Function getDescrizione()
		getDescrizione = descrizione
	End Function
	
	Public Sub setDescrizione(strDesc)
		descrizione = strDesc
	End Sub

	
	Public Function getValue()
		getValue = str_value
	End Function
	
	Public Sub setValue(strValue)
		str_value = strValue
	End Sub	
	
	Public Function getAlert()
		getAlert = alert
	End Function
	
	Public Sub setAlert(strAlert)
		alert = strAlert
	End Sub
	
	Public Function getType()
		getType = str_type
	End Function
	
	Public Sub setType(strType)
		str_type = strType
	End Sub
		
	Public Function getListaConfig()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objConfig
		getListaConfig = null		
		strSQL = "SELECT * FROM config_portal ORDER BY tipo, keyword;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF			
				Set objConfig = New ConfigClass
				strKey = objRS("keyword")
				strDesc = objRS("descrizione")	
				strValue = objRS("conf_value")
				strAlert = objRS("alert")
				strType = objRS("tipo")
				objConfig.setKey(strKey)	
				objConfig.setDescrizione(strDesc)
				objConfig.setValue(strValue)	
				objConfig.setAlert(strAlert)	
				objConfig.setType(strType)	
				objDict.add strKey, objConfig
				Set objConfig = nothing
				objRS.moveNext()
			loop
							
			Set getListaConfig = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getConfigPerKey(key)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getConfigPerKey = null		
		strSQL = "SELECT * FROM config_portal  WHERE keyword=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,key)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objConfig = New ConfigClass
			
			do while not objRS.EOF
				strKey = objRS("keyword")
				strDesc = objRS("descrizione")	
				strValue = objRS("conf_value")
				strAlert = objRS("alert")
				strType = objRS("tipo")
				objConfig.setKey(strKey)	
				objConfig.setDescrizione(strDesc)
				objConfig.setValue(strValue)	
				objConfig.setAlert(strAlert)	
				objConfig.setType(strType)	
				objDict.add strKey, objConfig
				objRS.moveNext()
			loop
							
			Set getConfigPerKey = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

			
	Public Sub insertConfig(strKey, strDescrizione, strValue, strAlert, strType)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO config_portal(keyword, descrizione, conf_value, alert, tipo) VALUES("
		strSQL = strSQL & "?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,strAlert)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strType)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyConfig(strKey, strDescrizione, strValue, strAlert, strType)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE config_portal SET "
		strSQL = strSQL & "keyword=?,"
		strSQL = strSQL & "descrizione=?,"
		strSQL = strSQL & "conf_value=?,"
		strSQL = strSQL & "alert=?,"
		strSQL = strSQL & "tipo=?"
		strSQL = strSQL & " WHERE key=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,129,1,1,strAlert)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strType)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKey)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub updateConfigValue(strKey, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE config_portal SET "
		strSQL = strSQL & "conf_value=?"
		strSQL = strSQL & " WHERE keyword=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKey)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
			
	Public Sub deleteConfig(strKey)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM config_portal WHERE keyword=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strKey)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
			
	Public Sub setAllApplicationVariables()
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "SELECT * FROM config_portal;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Set objRS = objConn.Execute(strSQL)
				
		if not(objRS.EOF) then
			do while not objRS.EOF
				strKey = objRS("keyword")
				Application(strKey) = objRS("conf_value")
				objRS.moveNext()			
			loop
		end if
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
End Class
%>