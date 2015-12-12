<%

Class CountryClass
	Private id
	Private country_code
	Private state_region_code
	Private country_description
	Private state_region_description
	Private active
	Private use_for
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub
	
	Public Function getCountryCode()
		getCountryCode = country_code
	End Function
	
	Public Sub setCountryCode(strCCode)
		country_code = strCCode
	End Sub
	
	Public Function getStateRegionCode()
		getStateRegionCode = state_region_code
	End Function
	
	Public Sub setStateRegionCode(strCRCode)
		state_region_code = strCRCode
	End Sub
	
	Public Function getCountryDescription()
		getCountryDescription = country_description
	End Function
	
	Public Sub setCountryDescription(strDesc)
		country_description = strDesc
	End Sub
	
	Public Function getStateRegionDescription()
		getStateRegionDescription = state_region_description
	End Function
	
	Public Sub setStateRegionDescription(strSRDesc)
		state_region_description = strSRDesc
	End Sub
	
	Public Sub setActive(strActive)
		active = strActive
	End Sub
	
	Public Function isActive()
		isActive = active
	End Function
	
	Public Function getUseFor()
		getUseFor = use_for
	End Function
	
	Public Sub setUseFor(strUseFor)
		use_for = strUseFor
	End Sub
	
		
	Public Function getListaCountry(active, useFor, countryCode, stateRegionCode)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaCountry = null		
		strSQL = "SELECT * FROM countries"

		if (isNull(active) AND isNull(useFor) AND isNull(countryCode) AND isNull(stateRegionCode)) then
			strSQL = "SELECT * FROM countries"
		else				
			strSQL = strSQL & " WHERE"

			if not(isNull(active)) then strSQL = strSQL & " AND active=?"
			if not(isNull(useFor)) then
				arrUseFor = Split(useFor, ",", -1, 1)
				if(Ubound(arrUseFor) > 0) then
					strSQL = strSQL & " AND("
					for each e in arrUseFor
						strSQL = strSQL & " use_for=? OR"
					next
					strSQL = strSQL & ")"
					strSQL = Replace(strSQL, " OR)", ")", 1, -1, 1)
				end if
			end if
			if not(isNull(countryCode)) then strSQL = strSQL & " AND country_code=?"
			if not(isNull(stateRegionCode)) then strSQL = strSQL & " AND state_region_code=?"
		end if
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & " ORDER BY country_description, state_region_description;"


		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		if (isNull(active) AND isNull(useFor) AND isNull(countryCode) AND isNull(stateRegionCode)) then
		else	
			if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active) end if
			if not(isNull(useFor)) then
				arrUseFor = Split(useFor, ",", -1, 1)
				if(Ubound(arrUseFor) > 0) then
					for each e in arrUseFor
						objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,e)
					next
				end if
			end if
			if not(isNull(countryCode)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode) end if
			if not(isNull(stateRegionCode)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode) end if
		end if
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strCountryCode = objRS("country_code")
				strStateRegionCode = objRS("state_region_code")
				strDesc = objRS("country_description")
				strSRDesc = objRS("state_region_description")
				bolActive = objRS("active")
				strUserFor = objRS("use_for")
						
				Set objCountryTmp = new CountryClass
				objCountryTmp.setID(strID)
				objCountryTmp.setCountryCode(strCountryCode)
				objCountryTmp.setStateRegionCode(strStateRegionCode)
				objCountryTmp.setCountryDescription(strDesc)	
				objCountryTmp.setStateRegionDescription(strSRDesc)
				objCountryTmp.setActive(bolActive)		
				objCountryTmp.setUseFor(strUserFor)									
				objDict.add strID, objCountryTmp
				Set objCountryTmp = Nothing
				objRS.moveNext()
			loop
							
			Set getListaCountry = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function findCountry(srtSearch)
		on error resume next
		findCountry = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT * FROM countries"

		if not(isNull(srtSearch)) then				
			strSQL = strSQL & " WHERE country_code LIKE ? OR state_region_code LIKE ? OR country_description LIKE ? OR state_region_description LIKE ?"
		end if
		strSQL = strSQL & " ORDER BY country_description, state_region_description;"
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		if not(isNull(srtSearch)) then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&srtSearch&"%")
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&srtSearch&"%")
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&srtSearch&"%")
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&srtSearch&"%")
		end if

		'****** METODI DI UTILITA' CONTROLLO COLLECTION PARAMETERS DI ADODB.COMMAND
		'response.write("typename(objCommand): "&typename(objCommand)&"<br>")
		'response.write("objCommand.State: "&objCommand.State&"<br>")
		'response.write("typename(objCommand.Parameters): "&typename(objCommand.Parameters)&" - size: "&objCommand.Parameters.Count&"<br>")
		'for each j in objCommand.Parameters
		'	response.write("j: "&j&" - type: "&j.type&" - size: "&j.size&" - direction: "&j.direction&" - value: "&j.value&"<br>")
		'next

		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strCountryCode = objRS("country_code")
				strStateRegionCode = objRS("state_region_code")
				strDesc = objRS("country_description")
				strSRDesc = objRS("state_region_description")
				bolActive = objRS("active")
				strUserFor = objRS("use_for")
						
				Set objCountryTmp = new CountryClass
				objCountryTmp.setID(strID)
				objCountryTmp.setCountryCode(strCountryCode)
				objCountryTmp.setStateRegionCode(strStateRegionCode)
				objCountryTmp.setCountryDescription(strDesc)	
				objCountryTmp.setStateRegionDescription(strSRDesc)
				objCountryTmp.setActive(bolActive)		
				objCountryTmp.setUseFor(strUserFor)									
				objDict.add strID, objCountryTmp
				Set objCountryTmp = Nothing
				objRS.moveNext()
			loop
							
			Set findCountry = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findCountryListOnly(useFor)
		on error resume next
		findCountryListOnly = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT * FROM countries WHERE active=1 AND (ISNULL(state_region_code) OR state_region_code='')"

		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				strSQL = strSQL & " AND("
				for each e in arrUseFor
					strSQL = strSQL & " use_for=? OR"
				next
				strSQL = strSQL & ")"
				strSQL = Replace(strSQL, " OR)", ")", 1, -1, 1)
			end if
		end if

		strSQL = Trim(strSQL)
		strSQL = strSQL & " ORDER BY country_description;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				for each e in arrUseFor
					objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,e)
				next
			end if
		end if
		Set objRS = objCommand.Execute()
	
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")

			do while not objRS.EOF
				strID = objRS("id")
				strCountryCode = objRS("country_code")
				strStateRegionCode = objRS("state_region_code")
				strDesc = objRS("country_description")
				strSRDesc = objRS("state_region_description")
				bolActive = objRS("active")
				strUserFor = objRS("use_for")
						
				Set objCountryTmp = new CountryClass
				objCountryTmp.setID(strID)
				objCountryTmp.setCountryCode(strCountryCode)
				objCountryTmp.setStateRegionCode(strStateRegionCode)
				objCountryTmp.setCountryDescription(strDesc)	
				objCountryTmp.setStateRegionDescription(strSRDesc)
				objCountryTmp.setActive(bolActive)		
				objCountryTmp.setUseFor(strUserFor)									
				objDict.add strID, objCountryTmp
				Set objCountryTmp = Nothing
				objRS.moveNext()
			loop
					
			Set findCountryListOnly = objDict			
			Set objDict = nothing
		end if

		
		Set objRS = Nothing
		Set objCommand = Nothing	
		Set objDB = Nothing	
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findStateRegionListOnly(useFor)
		on error resume next
		findStateRegionListOnly = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT * FROM countries WHERE active=1 AND (NOT ISNULL(state_region_code) AND state_region_code <>'')"

		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				strSQL = strSQL & " AND("
				for each e in arrUseFor
					strSQL = strSQL & " use_for=? OR"
				next
				strSQL = strSQL & ")"
				strSQL = Replace(strSQL, " OR)", ")", 1, -1, 1)
			end if
		end if

		strSQL = Trim(strSQL)
		strSQL = strSQL & " ORDER BY country_description, state_region_description;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
	
		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				for each e in arrUseFor
					objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,e)
				next
			end if
		end if

		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strCountryCode = objRS("country_code")
				strStateRegionCode = objRS("state_region_code")
				strDesc = objRS("country_description")
				strSRDesc = objRS("state_region_description")
				bolActive = objRS("active")
				strUserFor = objRS("use_for")
						
				Set objCountryTmp = new CountryClass
				objCountryTmp.setID(strID)
				objCountryTmp.setCountryCode(strCountryCode)
				objCountryTmp.setStateRegionCode(strStateRegionCode)
				objCountryTmp.setCountryDescription(strDesc)	
				objCountryTmp.setStateRegionDescription(strSRDesc)
				objCountryTmp.setActive(bolActive)		
				objCountryTmp.setUseFor(strUserFor)									
				objDict.add strID, objCountryTmp
				Set objCountryTmp = Nothing
				objRS.moveNext()
			loop
							
			Set findStateRegionListOnly = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findStateRegionListByCountry(countryCode, useFor)
		on error resume next
		findStateRegionListByCountry = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT * FROM countries WHERE active=1 AND (NOT ISNULL(state_region_code) AND state_region_code <>'')"

		if (not(isNull(countryCode)) AND countryCode<>"") then strSQL = strSQL & " AND country_code=?"

		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				strSQL = strSQL & " AND("
				for each e in arrUseFor
					strSQL = strSQL & " use_for=? OR"
				next
				strSQL = strSQL & ")"
				strSQL = Replace(strSQL, " OR)", ")", 1, -1, 1)
			end if
		end if

		strSQL = Trim(strSQL)
		strSQL = strSQL & " ORDER BY country_description, state_region_description;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
	
		if not(isNull(countryCode)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode) end if
		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				for each e in arrUseFor
					objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,e)
				next
			end if
		end if

		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strCountryCode = objRS("country_code")
				strStateRegionCode = objRS("state_region_code")
				strDesc = objRS("country_description")
				strSRDesc = objRS("state_region_description")
				bolActive = objRS("active")
				strUserFor = objRS("use_for")
						
				Set objCountryTmp = new CountryClass
				objCountryTmp.setID(strID)
				objCountryTmp.setCountryCode(strCountryCode)
				objCountryTmp.setStateRegionCode(strStateRegionCode)
				objCountryTmp.setCountryDescription(strDesc)	
				objCountryTmp.setStateRegionDescription(strSRDesc)
				objCountryTmp.setActive(bolActive)		
				objCountryTmp.setUseFor(strUserFor)									
				objDict.add strID, objCountryTmp
				Set objCountryTmp = Nothing
				objRS.moveNext()
			loop
							
			Set findStateRegionListByCountry = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	
		
	Public Function findCountryListCodeDesc(useFor)
		on error resume next
		findCountryListCodeDesc = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT country_code, country_description FROM countries WHERE active=1 AND (ISNULL(state_region_code) OR state_region_code='')"

		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				strSQL = strSQL & " AND("
				for each e in arrUseFor
					strSQL = strSQL & " use_for=? OR"
				next
				strSQL = strSQL & ")"
				strSQL = Replace(strSQL, " OR)", ")", 1, -1, 1)
			end if
		end if

		strSQL = Trim(strSQL)
		strSQL = strSQL & " ORDER BY country_code;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				for each e in arrUseFor
					objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,e)
				next
			end if
		end if
		Set objRS = objCommand.Execute()
	
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")

			do while not objRS.EOF
				strCountryCode = objRS("country_code")
				strDesc = objRS("country_description")							
				objDict.add strCountryCode, strDesc
				objRS.moveNext()
			loop
					
			Set findCountryListCodeDesc = objDict			
			Set objDict = nothing
		end if

		
		Set objRS = Nothing
		Set objCommand = Nothing	
		Set objDB = Nothing	
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findStateRegionListCodeDesc(useFor)
		on error resume next
		findStateRegionListCodeDesc = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT state_region_code, state_region_description FROM countries WHERE active=1 AND (NOT ISNULL(state_region_code) AND state_region_code <>'')"

		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				strSQL = strSQL & " AND("
				for each e in arrUseFor
					strSQL = strSQL & " use_for=? OR"
				next
				strSQL = strSQL & ")"
				strSQL = Replace(strSQL, " OR)", ")", 1, -1, 1)
			end if
		end if

		strSQL = Trim(strSQL)
		strSQL = strSQL & " ORDER BY state_region_code;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		if not(isNull(useFor)) then
			arrUseFor = Split(useFor, ",", -1, 1)
			if(Ubound(arrUseFor) > 0) then
				for each e in arrUseFor
					objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,e)
				next
			end if
		end if

		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strStateRegionCode = objRS("state_region_code")
				strSRDesc = objRS("state_region_description")													
				objDict.add strStateRegionCode, strSRDesc
				objRS.moveNext()
			loop
							
			Set findStateRegionListCodeDesc = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	
		
	Public Function findCountryByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findCountryByID = null		
		strSQL = "SELECT * FROM countries WHERE id=?;"	
		
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
			Dim objCountryTmp
			
			strID = objRS("id")
			strCountryCode = objRS("country_code")
			strStateRegionCode = objRS("state_region_code")
			strDesc = objRS("country_description")
			strSRDesc = objRS("state_region_description")
			bolActive = objRS("active")
			strUserFor = objRS("use_for")
				
			Set objCountryTmp = new CountryClass
			objCountryTmp.setID(strID)
			objCountryTmp.setCountryCode(strCountryCode)
			objCountryTmp.setStateRegionCode(strStateRegionCode)
			objCountryTmp.setCountryDescription(strDesc)	
			objCountryTmp.setStateRegionDescription(strSRDesc)
			objCountryTmp.setActive(bolActive)		
			objCountryTmp.setUseFor(strUserFor)			
			Set findCountryByID = objCountryTmp
			Set objCountryTmp = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
				
	Public Function insertCountry(countryCode, strDesc, stateRegionCode, strSRDesc, bolActive, strUserFor)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		insertCountry = -1
		
		strSQL = "INSERT INTO countries(country_code, country_description, state_region_code, state_region_description,active,use_for) VALUES("
		strSQL = strSQL & "?,?,"
		if(isNull(stateRegionCode) OR stateRegionCode = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"			
		end if
		if(isNull(strSRDesc) OR strSRDesc = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"			
		end if		
		strSQL = strSQL & "?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strDesc)
		if not isNull(stateRegionCode) AND not(stateRegionCode = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode)
		end if
		if not isNull(strSRDesc) AND not(strSRDesc = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strSRDesc)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strUserFor)
		objCommand.Execute()		
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(countries.id) as id FROM countries")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertCountry = objRS("id")	
		end if		
		Set objRS = Nothing
		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function	
				
	Public Sub modifyCountry(id,countryCode, strDesc, stateRegionCode, strSRDesc, bolActive, strUserFor)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "UPDATE countries SET "
		strSQL = strSQL & "country_code=?,"
		strSQL = strSQL & "country_description=?,"
		if(isNull(stateRegionCode) OR stateRegionCode = "") then
			strSQL = strSQL & "state_region_code=NULL,"
		else
			strSQL = strSQL & "state_region_code=?,"			
		end if
		if(isNull(strSRDesc) OR strSRDesc = "") then
			strSQL = strSQL & "state_region_description=NULL,"
		else
			strSQL = strSQL & "state_region_description=?,"			
		end if
		strSQL = strSQL & "`active`=?,"		
		strSQL = strSQL & "use_for=?"
		strSQL = strSQL & " WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strDesc)
		if not isNull(stateRegionCode) AND not(stateRegionCode = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode)
		end if
		if not isNull(strSRDesc) AND not(strSRDesc = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strSRDesc)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolActive)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strUserFor)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteCountry(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM countries WHERE id=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

End Class
%>