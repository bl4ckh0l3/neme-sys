<%

Class LogClass
	
	Private id
	Private msg
	Private usr
	Private tipo	
	Private data
	
	
	Public Function getLogID()
		getLogID = id
	End Function
	
	Public Sub setLogID(strID)
		id = strID
	End Sub	
	
	Public Function getLogMsg()
		getLogMsg = msg
	End Function
	
	Public Sub setLogMsg(strMsg)
		msg = strMsg
	End Sub	
	
	
	Public Function getLogUsr()
		getLogUsr = usr
	End Function
	
	Public Sub setLogUsr(strUsr)
		usr = strUsr
	End Sub
	
	
	Public Function getLogTipo()
		getLogTipo = tipo
	End Function
	
	Public Sub setLogTipo(strTipo)
		tipo = strTipo
	End Sub
	
	
	Public Function getLogData()
		getLogData = data
	End Function
	
	Public Sub setLogData(strData)
		data = strData
	End Sub	
	
	
	Public Function getListaLogs(tipo, dta_from, dta_to)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		Dim DD, MM, YY, HH, MIN, SS

		getListaLogs = null		
		strSQL = "SELECT * FROM logs"
		
		if ((isNull(tipo) OR tipo="") AND (isNull(dta_from) OR dta_from="") AND (isNull(dta_to) OR dta_to="")) then
			strSQL = "SELECT * FROM logs"
		else
			strSQL = strSQL & " WHERE"
			if not(isNull(tipo)) AND tipo<>"" then strSQL = strSQL & " AND type=?"
			'il passaggio seguente è da verificare con query secca di test su DB
			if not(isNull(dta_from)) AND dta_from<>"" then 		
				DD = DatePart("d", dta_from)
				MM = DatePart("m", dta_from)
				YY = DatePart("yyyy", dta_from)
				HH = 00
				MIN = 00
				SS = 00
				dta_from = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS

				strSQL = strSQL & " AND date_event >=?"
			end if
			
			if not(isNull(dta_to)) AND dta_to<>"" then 
				DD = DatePart("d", dta_to)
				MM = DatePart("m", dta_to)
				YY = DatePart("yyyy", dta_to)
				HH = 23
				MIN = 59
				SS = 59
				dta_to = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS
				
				strSQL = strSQL & " AND date_event <=?"
			end if			
		end if
		
		strSQL = strSQL & " ORDER BY date_event DESC;"
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if ((isNull(tipo) OR tipo="") AND (isNull(dta_from) OR dta_from="") AND (isNull(dta_to) OR dta_to="")) then
		else
			if not(isNull(tipo)) AND tipo<>"" then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,15,tipo)
			if not(isNull(dta_from)) AND dta_from<>"" then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_from)
			if not(isNull(dta_to)) AND dta_to<>"" then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_to)		
		end if
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objLog
			do while not objRS.EOF				
				Set objLog = new LogClass
				strID = objRS("id")
				objLog.setLogID(strID)
				objLog.setLogMsg(objRS("msg"))
				objLog.setLogUsr(objRS("usr"))	
				objLog.setLogTipo(objRS("type"))	
				objLog.setLogData(objRS("date_event"))	
				objDict.add strID, objLog
				objRS.moveNext()
			loop
			Set objLog = nothing							
			Set getListaLogs = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	

	Public Sub write(strMsg, usr, tipo)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		Dim dta_ins, DD, MM, YY, HH, MIN, SS
		
		dta_ins = Now()
		strSQL = "INSERT INTO logs(msg, usr, type, date_event) VALUES(?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMsg)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,usr)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,15,tipo)
		objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_ins)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
	
	Public Sub deleteLogs(tipo, dta_from, dta_to)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		Dim DD, MM, YY, HH, MIN, SS		
		
		strSQL = "DELETE FROM logs"
		
		if ((isNull(tipo) OR tipo="") AND (isNull(dta_from) OR dta_from="") AND (isNull(dta_to) OR dta_to="")) then
			strSQL = "DELETE FROM logs;"
		else
			strSQL = strSQL & " WHERE"
			
			if not(isNull(tipo)) AND tipo<>"" then strSQL = strSQL & " AND type=?"
			'il passaggio seguente è da verificare con query secca di test su DB
			if not(isNull(dta_from)) AND dta_from<>"" then 		
				DD = DatePart("d", dta_from)
				MM = DatePart("m", dta_from)
				YY = DatePart("yyyy", dta_from)
				HH = 00
				MIN = 00
				SS = 00
				dta_from = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS

				strSQL = strSQL & " AND date_event >=?"
			end if

			if not(isNull(dta_to)) AND dta_to<>"" then 
				DD = DatePart("d", dta_to)
				MM = DatePart("m", dta_to)
				YY = DatePart("yyyy", dta_to)
				HH = 23
				MIN = 59
				SS = 59
				dta_to = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS
				
				strSQL = strSQL & " AND date_event <=?"
			end if
		end if
		
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if ((isNull(tipo) OR tipo="") AND (isNull(dta_from) OR dta_from="") AND (isNull(dta_to) OR dta_to="")) then
		else
			if not(isNull(tipo)) AND tipo<>"" then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,15,tipo)
			if not(isNull(dta_from)) AND dta_from<>"" then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_from)
			if not(isNull(dta_to)) AND dta_to<>"" then objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,dta_to)			
		end if
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Sub

End Class
%>