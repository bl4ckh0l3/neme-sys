<%
Class TemplateClass
	Private id
	Private templateCss
	Private descrizioneTemplate
	Private baseTemplate
	Private dir_template
	Private orderBy
	Private elemXpage
	Private objPagePerTemplate
	
	Public Function getID()
		getID = id
	End Function
				
	Public Sub setID(numID)
		id = numID
	End Sub
		
	Public Function getDirTemplate()
		getDirTemplate = dir_template
	End Function
				
	Public Sub setDirTemplate(strDirTemplate)
		dir_template = strDirTemplate
	End Sub	
		
	Public Function getTemplateCssWithPath()	
		if not(templateCss = "") then
			getTemplateCssWithPath = dir_template&"/css/"&templateCss
		else
			getTemplateCssWithPath = templateCss		
		end if
	End Function
		
	Public Function getTemplateCss()
		getTemplateCss = templateCss
	End Function
				
	Public Sub setTemplateCss(strTemplateCss)
		templateCss = strTemplateCss
	End Sub
	
	Public Function getDescrizioneTemplate()
		getDescrizioneTemplate = descrizioneTemplate
	End Function
				
	Public Sub setDescrizioneTemplate(strDescrizioneTemplate)
		descrizioneTemplate = strDescrizioneTemplate
	End Sub
	
	Public Function getBaseTemplate()
		getBaseTemplate = baseTemplate
	End Function
				
	Public Sub setBaseTemplate(numBaseTemplate)
		baseTemplate = numBaseTemplate
	End Sub
	
	Public Function getOrderBy()
		getOrderBy = orderBy
	End Function
				
	Public Sub setOrderBy(numOrderBy)
		orderBy = numOrderBy
	End Sub
	
	Public Function getElemXPage()
		getElemXPage = elemXpage
	End Function
				
	Public Sub setElemXPage(numElemXpage)
		elemXpage = numElemXpage
	End Sub
	
	Public Function getPagePerTemplate()		
		if(isNull(objPagePerTemplate)) then
			getPagePerTemplate = null
		else
			Set getPagePerTemplate = objPagePerTemplate
		end if
	End Function
	
	Public Sub setPagePerTemplate(objPage)
		if(isNull(objPage)) then
			objPagePerTemplate = null
		else
			Set objPagePerTemplate = objPage
		end if		
	End Sub			


'*********************************** METODI TEMPLATE *********************** 				
	Public Function insertTemplate(strDirTemplate, strTemplateCss, strDescrizioneTemplate, baseTemplate, orderBy, elemXpage, objConn)
		on error resume next
		insertTemplate = -1
		
		Dim strSQL, strSQLSelect, objRS
		
		strSQL = "INSERT INTO template_disponibili(dir_template, template_css, descrizione, base_template, order_by, elem_x_page) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strDirTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTemplateCss)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizioneTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,baseTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,orderBy)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,elemXpage)
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(id) as id FROM template_disponibili")
		if not (objRS.EOF) then
			insertTemplate = objRS("id")	
		end if		
		Set objRS = Nothing
				
		if objConn.Errors.Count <> 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyTemplate(id_template, strDirTemplate, strTemplateCss, strDescrizioneTemplate, baseTemplate, orderBy, elemXpage, objConn)
		on error resume next
		Dim strSQL, objRS
		strSQL = "UPDATE template_disponibili SET "
		strSQL = strSQL & "dir_template=?,"
		strSQL = strSQL & "template_css=?,"
		strSQL = strSQL & "descrizione=?,"
		strSQL = strSQL & "base_template=?,"
		strSQL = strSQL & "order_by=?,"
		strSQL = strSQL & "elem_x_page=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,strDirTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTemplateCss)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizioneTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,baseTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,orderBy)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,elemXpage)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_template)
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
		
	Public Sub deleteTemplate(id)
		on error resume next
		Dim objDB, strSQLDel, strSQLDelPages, objRS, objConn
		strSQLDel = "DELETE FROM template_disponibili WHERE id=?;"
		strSQLDelPages = "DELETE FROM page_x_template WHERE id_template=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		Dim objCommand, objCommand2
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQLDel
		objCommand2.CommandText = strSQLDelPages
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		
		objConn.BeginTrans
		
		objCommand2.Execute()
		objCommand.Execute()
	
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
		
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Function getListaTemplates()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objTemplate
		getListaTemplates = null  
		strSQL = "SELECT * FROM template_disponibili ORDER BY descrizione;"
		
		Set objDB = New DBManagerClass
		Set objPagePerTemplate = new Page4TemplateClass
		Set objConn = objDB.openConnection()  
		Set objRS = objConn.Execute(strSQL)
		
		if not(objRS.EOF) then		   
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
			do while not objRS.EOF
				Set objTemplate = new TemplateClass	
				strID = objRS("id")				
				objTemplate.setID(strID)    
				objTemplate.setDirTemplate(objRS("dir_template"))
				objTemplate.setTemplateCss(objRS("template_css"))
				objTemplate.setDescrizioneTemplate(objRS("descrizione"))		
				objTemplate.setBaseTemplate(objRS("base_template"))
				objTemplate.setOrderBy(objRS("order_by"))
				objTemplate.setElemXPage(objRS("elem_x_page"))
				
				Set objPage = objPagePerTemplate.getListaPagePerTemplates(strID, false)				
				if not(isEmpty(objPage)) then
					objTemplate.setPagePerTemplate(objPage)
					Set objPage = nothing
				else
					Set objPage = nothing
				end if	
													
				objDict.add strID, objTemplate
				Set objTemplate = Nothing
				objRS.moveNext()
			loop
			
			Set getListaTemplates = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		Set objPagePerTemplate = nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function
				
	Public Function findTemplateByID(id)
		on error resume next
		
		findTemplateByID = null
		
		Dim objDB, strSQL, objRS, objConn, strID
		strSQL = "SELECT * FROM template_disponibili WHERE id=?;"
		
		Set objPagePerTemplate = new Page4TemplateClass
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
			Dim objTemplate
			Set objTemplate = new TemplateClass	
			strID = objRS("id")
			objTemplate.setID(strID)    
			objTemplate.setDirTemplate(objRS("dir_template"))
			objTemplate.setTemplateCss(objRS("template_css"))
			objTemplate.setDescrizioneTemplate(objRS("descrizione"))	
			objTemplate.setBaseTemplate(objRS("base_template"))
			objTemplate.setOrderBy(objRS("order_by"))
			objTemplate.setElemXPage(objRS("elem_x_page"))

			Set objPage = objPagePerTemplate.getListaPagePerTemplates(strID, false)				
			if not(isEmpty(objPage)) then
				objTemplate.setPagePerTemplate(objPage)
				Set objPage = nothing
			else
				Set objPage = nothing
			end if	
						
			Set findTemplateByID = objTemplate
			Set objTemplate = Nothing
		end if	
	
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		Set objPagePerTemplate = nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getMaxIDTemplate()
		on error resume next
		
		getMaxIDTemplate = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT MAX(id) AS id_template FROM template_disponibili;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDTemplate = objRS("id_template")	
		end if
				
		Set objRS = Nothing
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