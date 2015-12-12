<%
Class Page4TemplateClass
	Private id
	Private idTemplate
	Private file_name
	Private page_num	
	
	Public Function getID()
		getID = id
	End Function
				
	Public Sub setID(numID)
		id = numID
	End Sub
	
	Public Function getIDTemplate()
		getIDTemplate = idTemplate
	End Function
				
	Public Sub setIDTemplate(numIDTemplate)
		idTemplate = numIDTemplate
	End Sub
		
	Public Function getFileName()
		getFileName = file_name
	End Function
				
	Public Sub setFileName(strFileName)
		file_name = strFileName
	End Sub	
		
	Public Function getPageNum()
		getPageNum = page_num
	End Function
				
	Public Sub setPageNum(strPageNum)
		page_num = strPageNum
	End Sub


'*********************************** METODI TEMPLATE *********************** 				
	Public Sub insertPagePerTemplate(strIDTemplate, strFile, pageNum, objConn)
		on error resume next
		Dim strSQL, strSQLSelect, objRS
		
		strSQL = "INSERT INTO page_x_template(id_template, file_name, page_num) VALUES("
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strIDTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strFile)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,pageNum)
		objCommand.Execute()
		Set objCommand = Nothing
		
		if objConn.Errors.Count <> 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyPagePerTemplate(id, strIDTemplate, strFile, pageNum, objConn)
		on error resume next
		Dim strSQL, objRS
		strSQL = "UPDATE page_x_template SET "
		strSQL = strSQL & "id_template=?,"
		strSQL = strSQL & "file_name=?,"
		strSQL = strSQL & "page_num=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strIDTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strFile)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,pageNum)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		
		if objConn.Errors.Count <> 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if			
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deletePagePerTemplate(id)
		on error resume next
		Dim objDB, strSQLDel, strSQLDelPages, objRS, objConn
		strSQLDelPages = "DELETE FROM page_x_template WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDelPages
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Function getListaPagePerTemplates(id_template, bolNotInclude)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objPageXTemplate
		getListaPagePerTemplates = null  
		strSQL = "SELECT * FROM page_x_template WHERE id_template=?"
		if(bolNotInclude)then
			strSQL = strSQL & " AND page_num <> -1"
		end if
		strSQL = strSQL & " ORDER BY page_num, file_name;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_template)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then		   
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
			do while not objRS.EOF
				Set objPageXTemplate = new Page4TemplateClass	
				strID = objRS("id")				
				objPageXTemplate.setID(strID)    
				objPageXTemplate.setIDTemplate(objRS("id_template"))
				objPageXTemplate.setFileName(objRS("file_name"))
				objPageXTemplate.setPageNum(objRS("page_num"))									
				objDict.add strID, objPageXTemplate
				Set objPageXTemplate = Nothing
				objRS.moveNext()
			loop
			
			Set getListaPagePerTemplates = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
				
	Public Function findPagePerTemplateByID(id)
		on error resume next
		
		findPagePerTemplateByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM page_x_template WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		Set objRS = objCommand.Execute()		
		
		if not objRS.EOF then
			Dim objPageXTemplate
			Set objPageXTemplate = new Page4TemplateClass	
			strID = objRS("id")				
			objPageXTemplate.setID(strID)    
			objPageXTemplate.setIDTemplate(objRS("id_template"))
			objPageXTemplate.setFileName(objRS("file_name"))
			objPageXTemplate.setPageNum(objRS("page_num"))			
			Set findPagePerTemplateByID = objPageXTemplate
			Set objPageXTemplate = Nothing	
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
				
	Public Function findPageByNum(idTemplate, numPage)
		on error resume next
		
		findPageByNum = null
		
		Dim objDB, strSQL, objRS, objConn, strID
		strSQL = "SELECT * FROM page_x_template WHERE id_template=? AND page_num=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numPage)
		Set objRS = objCommand.Execute()		

		if not objRS.EOF then		
			findPageByNum = objRS("file_name")
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
				
	Public Function findPagePerTemplateByFileName(id_template, filename)
		on error resume next
		
		findPagePerTemplateByFileName = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM page_x_template WHERE id_template=? AND file_name=?"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_template)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,filename)
		Set objRS = objCommand.Execute()		
		
		if not objRS.EOF then
			Dim objPageXTemplate
			Set objPageXTemplate = new Page4TemplateClass	
			strID = objRS("id")				
			objPageXTemplate.setID(strID)    
			objPageXTemplate.setIDTemplate(objRS("id_template"))
			objPageXTemplate.setFileName(objRS("file_name"))
			objPageXTemplate.setPageNum(objRS("page_num"))
			Set findPagePerTemplateByFileName = objPageXTemplate
			Set objPageXTemplate = Nothing				
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getMaxIDPagePerTemplate()
		on error resume next
		
		getMaxIDPagePerTemplate = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT MAX(id) AS id FROM page_x_template;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDPagePerTemplate = objRS("id")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getMaxNumPageByIDTemplate(id_template)
		on error resume next
		
		getMaxNumPageByIDTemplate = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT MAX(page_num) AS page_num FROM page_x_template WHERE id_template=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_template)
		Set objRS = objCommand.Execute()				
		
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxNumPageByIDTemplate = objRS("page_num")	
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	'public Sub toString()
		'response.write ()
	'end Sub
End Class
%>