<%
Class ContentFieldClass
	Private id
	Private description
	Private idGroup
	Private objGroup
	Private order
	Private typeField
	Private typeContent
	Private maxLenght
	Private required
	Private enabled
	Private editable
	Private idContent
	Private selValue
	
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strid)
		id = strid
	End Sub	
	
	Public Function getDescription()
		getdescription = description
	End Function
	
	Public Sub setDescription(strDescription)
		description = strDescription
	End Sub
	
	Public Function getIdGroup()
		getIdGroup = idGroup
	End Function
	
	Public Sub setIdGroup(strIdGroup)
		idGroup = strIdGroup
	End Sub	
	
	Public Function getObjGroup()
		Set getObjGroup = objGroup
	End Function
	
	Public Sub setObjGroup(strObjGroup)
		Set objGroup = strObjGroup
	End Sub	
	
	Public Function getOrder()
		getOrder = order
	End Function
	
	Public Sub setOrder(strOrder)
		order = strOrder
	End Sub
	
	Public Function getTypeField()
		getTypeField = typeField
	End Function
	
	Public Sub setTypeField(strTypeField)
		typeField = strTypeField
	End Sub
	
	Public Function getTypeContent()
		getTypeContent = typeContent
	End Function
	
	Public Sub setTypeContent(strTypeContent)
		typeContent = strTypeContent
	End Sub	
	
	Public Function getMaxLenght()
		getMaxLenght = maxLenght
	End Function
	
	Public Sub setMaxLenght(strMaxLenght)
		maxLenght = strMaxLenght
	End Sub	
	
	Public Function getRequired()
		getRequired = required
	End Function
	
	Public Sub setRequired(bolRequired)
		required = bolRequired
	End Sub
	
	Public Function getEnabled()
		getEnabled = enabled
	End Function
	
	Public Sub setEnabled(bolEnabled)
		enabled = bolEnabled
	End Sub
	
	Public Function getEditable()
		getEditable = editable
	End Function
	
	Public Sub setEditable(bolEditable)
		editable = bolEditable
	End Sub
	
	Public Function getIdContent()
		getIdContent = idContent
	End Function
	
	Public Sub setIdContent(strIdContent)
		idContent = strIdContent
	End Sub
	
	Public Function getSelValue()
		getSelValue = selValue
	End Function
	
	Public Sub setSelValue(strSelValue)
		selValue = strSelValue
	End Sub
	
	
	'************************* GESTIONE CONTENT FIELDS *******************************
		
	Public Function getListContentField(enabled)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListContentField = null		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder FROM content_fields LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id"

		if not(isNull(enabled)) then
			strSQL = strSQL & " WHERE enabled=?"
		end if
		
		strSQL = strSQL & " ORDER BY gorder, content_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if not(isNull(enabled)) then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		end if
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objContentField
			do while not objRS.EOF				
				Set objContentField = new ContentFieldClass
				strID = objRS("id")
				objContentField.setID(strID)
				objContentField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objContentField.setIdGroup(strIdGroup)
				
				Set objGroup = new ContentFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objContentField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objContentField.setOrder(objRS("order"))	
				objContentField.setTypeField(objRS("type"))
				objContentField.setTypeContent(objRS("type_content"))
				objContentField.setMaxLenght(objRS("max_lenght"))	
				objContentField.setRequired(objRS("required"))	
				objContentField.setEnabled(objRS("enabled"))		
				objContentField.setEditable(objRS("editable"))	
				objDict.add strID, objContentField
				objRS.moveNext()
			loop
			Set objContentField = nothing							
			Set getListContentField = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function getListContentField4Content(idContent)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListContentField4Content = null		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder, content_fields_match.id_news, content_fields_match.value FROM content_fields"
		strSQL = strSQL & " LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id"
		strSQL = strSQL & " LEFT JOIN content_fields_match ON content_fields.id = content_fields_match.id_field AND content_fields_match.id_news=?"
		strSQL = strSQL & " WHERE enabled=1"			
		strSQL = strSQL & " ORDER BY gorder, content_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objContentField
			do while not objRS.EOF				
				Set objContentField = new ContentFieldClass
				strID = objRS("id")
				objContentField.setID(strID)
				objContentField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objContentField.setIdGroup(strIdGroup)
				
				Set objGroup = new ContentFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objContentField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objContentField.setOrder(objRS("order"))	
				objContentField.setTypeField(objRS("type"))
				objContentField.setTypeContent(objRS("type_content"))
				objContentField.setMaxLenght(objRS("max_lenght"))	
				objContentField.setRequired(objRS("required"))	
				objContentField.setEnabled(objRS("enabled"))
				objContentField.setEditable(objRS("editable"))	
				objContentField.setidContent(objRS("id_news"))
				objContentField.setSelValue(objRS("value"))
				objDict.add strID, objContentField
				objRS.moveNext()
			loop
			
			Set objContentField = nothing							
			Set getListContentField4Content = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListContentField4ContentActive(idContent)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListContentField4ContentActive = null		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder, content_fields_match.id_news, content_fields_match.value FROM content_fields"
		strSQL = strSQL & " LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id"
		strSQL = strSQL & " LEFT JOIN content_fields_match ON content_fields.id = content_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND content_fields_match.id_news=?"			
		strSQL = strSQL & " ORDER BY gorder, content_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objContentField
			do while not objRS.EOF				
				Set objContentField = new ContentFieldClass
				strID = objRS("id")
				objContentField.setID(strID)
				objContentField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objContentField.setIdGroup(strIdGroup)
				
				Set objGroup = new ContentFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objContentField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objContentField.setOrder(objRS("order"))	
				objContentField.setTypeField(objRS("type"))
				objContentField.setTypeContent(objRS("type_content"))
				objContentField.setMaxLenght(objRS("max_lenght"))		
				objContentField.setRequired(objRS("required"))	
				objContentField.setEnabled(objRS("enabled"))
				objContentField.setEditable(objRS("editable"))	
				objContentField.setidContent(objRS("id_news"))
				objContentField.setSelValue(objRS("value"))
				objDict.add strID, objContentField
				objRS.moveNext()
			loop
			
			Set objContentField = nothing							
			Set getListContentField4ContentActive = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListContentField4ContentActiveCached(idContent)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, cacheKey
		getListContentField4ContentActiveCached = null	

		cacheKey="listcf"&"-"&idContent
		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder, content_fields_match.id_news, content_fields_match.value FROM content_fields"
		strSQL = strSQL & " LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id"
		strSQL = strSQL & " LEFT JOIN content_fields_match ON content_fields.id = content_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND content_fields_match.id_news=?"			
		strSQL = strSQL & " ORDER BY gorder, content_fields.order;"


		'tento il recupero dell'oggetto dalla cache
		on error resume next
		Set ojbCache = new CacheClass
		
		Set cachedObj = ojbCache.getItem(cacheKey)
		
		if (Instr(1, typename(cachedObj), "Dictionary", 1) > 0) then
			Set objListaContentFieldC = Server.CreateObject("Scripting.Dictionary")
		
			for each skey in cachedObj
				Set objContentFieldC = new ContentFieldClass						
				
				objContentFieldC.setID(cachedObj(skey)("id"))
				objContentFieldC.setDescription(cachedObj(skey)("description"))
				objContentFieldC.setIdGroup(cachedObj(skey)("id_group"))
				
				Set objGroup = new ContentFieldGroupClass
				Set objCachedGroup = cachedObj(skey)("obj_group")
				objGroup.setID(objCachedGroup("id_group"))
				objGroup.setDescription(objCachedGroup("gdesc"))
				objGroup.setOrder(objCachedGroup("gorder"))		
				objContentFieldC.setObjGroup(objGroup)	
				Set objCachedGroup = nothing						
				Set objGroup = nothing						
				
				objContentFieldC.setOrder(cachedObj(skey)("order"))	
				objContentFieldC.setTypeField(cachedObj(skey)("type"))
				objContentFieldC.setTypeContent(cachedObj(skey)("type_content"))
				objContentFieldC.setMaxLenght(cachedObj(skey)("max_lenght"))		
				objContentFieldC.setRequired(cachedObj(skey)("required"))	
				objContentFieldC.setEnabled(cachedObj(skey)("enabled"))
				objContentFieldC.setEditable(cachedObj(skey)("editable"))	
				objContentFieldC.setidContent(cachedObj(skey)("id_news"))
				objContentFieldC.setSelValue(cachedObj(skey)("value"))					
			
				objListaContentFieldC.add skey, objContentFieldC
				Set objContentFieldC = nothing	
			next	
			
			Set getListContentField4ContentActiveCached = objListaContentFieldC
			Set objListaContentFieldC = nothing			
		else
			getListContentField4ContentActiveCached = null
		end if
		
		if Err.number <> 0 then
			getListContentField4ContentActiveCached = null
			'response.write(Err.number&" - "&Err.description&"<br>")
		end if
		

		if not(Instr(1, typename(getListContentField4ContentActiveCached), "Dictionary", 1) > 0) then

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()	
			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQL
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
			Set objRS = objCommand.Execute()
			
			if not(objRS.EOF) then			
				Set objDict = Server.CreateObject("Scripting.Dictionary")
				Set objDictCache = Server.CreateObject("Scripting.Dictionary")
				
				Dim objContentField
				do while not objRS.EOF				
					Set objContentField = new ContentFieldClass
					Set objContentFieldCache = Server.CreateObject("Scripting.Dictionary")
					
					strID = objRS("id")
					objContentField.setID(strID)
					objContentField.setDescription(objRS("description"))
					strIdGroup = objRS("id_group")
					objContentField.setIdGroup(strIdGroup)
					
					Set objGroup = new ContentFieldGroupClass
					objGroup.setID(strIdGroup)
					objGroup.setDescription(objRS("gdesc"))
					objGroup.setOrder(objRS("gorder"))		
					objContentField.setObjGroup(objGroup)		
					Set objGroup = nothing
					
					objContentField.setOrder(objRS("order"))	
					objContentField.setTypeField(objRS("type"))
					objContentField.setTypeContent(objRS("type_content"))
					objContentField.setMaxLenght(objRS("max_lenght"))		
					objContentField.setRequired(objRS("required"))	
					objContentField.setEnabled(objRS("enabled"))
					objContentField.setEditable(objRS("editable"))	
					objContentField.setidContent(objRS("id_news"))
					objContentField.setSelValue(objRS("value"))

					objContentFieldCache.add "id", strID
					objContentFieldCache.add "description", objContentField.getDescription()
					objContentFieldCache.add "id_group", objContentField.getIdGroup()
					
					Set objGroupC = Server.CreateObject("Scripting.Dictionary")
					objGroupC.add "id_group", objContentField.getIdGroup()
					objGroupC.add "gdesc", objContentField.getObjGroup().getDescription()
					objGroupC.add "gorder", objContentField.getObjGroup().getOrder()	
					objContentFieldCache.add "obj_group", objGroupC		
					Set objGroupC = nothing
					
					objContentFieldCache.add "order", objContentField.getOrder()
					objContentFieldCache.add "type", objContentField.getTypeField()
					objContentFieldCache.add "type_content", objContentField.getTypeContent()
					objContentFieldCache.add "max_lenght", objContentField.getMaxLenght()	
					objContentFieldCache.add "required", objContentField.getRequired()	
					objContentFieldCache.add "enabled", objContentField.getEnabled()
					objContentFieldCache.add "editable", objContentField.getEditable()	 
					objContentFieldCache.add "id_news", objContentField.getidContent()
					objContentFieldCache.add "value", objContentField.getSelValue()
					
					
					objDict.add strID, objContentField
					objDictCache.add strID, objContentFieldCache				
					Set objContentFieldCache = nothing
					objRS.moveNext()
				loop
				
				Set objContentField = nothing							
				Set getListContentField4ContentActiveCached = objDict	
				call ojbCache.store(cacheKey, objDictCache)	
				Set objDictCache = nothing				
				Set objDict = nothing				
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
		
	Public Function getListContentField4ContentActiveMultiple(listIdContents)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListContentField4ContentActiveMultiple = null		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder, content_fields_match.id_news, content_fields_match.value FROM content_fields"
		strSQL = strSQL & " LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id"
		strSQL = strSQL & " LEFT JOIN content_fields_match ON content_fields.id = content_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND content_fields_match.id_news IN("&listIdContents&")"			
		strSQL = strSQL & " ORDER BY id_news, gorder, content_fields.order;"
		'response.write(strSQL)
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		'objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objDictContentField = Server.CreateObject("Scripting.Dictionary")
			Dim objContentField
			idNews = -1
			do while not objRS.EOF				
				strID = objRS("id")
				idNewsTmp = objRS("id_news")
				
				Set objContentField = new ContentFieldClass
				objContentField.setID(strID)
				objContentField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objContentField.setIdGroup(strIdGroup)
				
				Set objGroup = new ContentFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objContentField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objContentField.setOrder(objRS("order"))	
				objContentField.setTypeField(objRS("type"))
				objContentField.setTypeContent(objRS("type_content"))
				objContentField.setMaxLenght(objRS("max_lenght"))		
				objContentField.setRequired(objRS("required"))	
				objContentField.setEnabled(objRS("enabled"))
				objContentField.setEditable(objRS("editable"))	
				objContentField.setidContent(idNewsTmp)
				objContentField.setSelValue(objRS("value"))
					
				if(idNewsTmp<>idNews)then
					objDict.add idNews, objDictContentField					
					Set objDictContentField = Server.CreateObject("Scripting.Dictionary")
					objDictContentField.add strID, objContentField
					idNews = idNewsTmp
				else					
					objDictContentField.add strID, objContentField
				end if				
				objRS.moveNext()
			loop
			
			Set objContentField = nothing							
			Set getListContentField4ContentActiveMultiple = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListContentField4ContentActiveByType(idContent, typeP)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListContentField4ContentActiveByType = null		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder, content_fields_match.id_news, content_fields_match.value FROM content_fields"
		strSQL = strSQL & " LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id"
		strSQL = strSQL & " LEFT JOIN content_fields_match ON content_fields.id = content_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND content_fields_match.id_news=?"	
		if not(isNull(typeP)) then
			strSQL = strSQL & " AND type IN("&typeP&")"	
		end if		
		strSQL = strSQL & " ORDER BY gorder, content_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objContentField
			do while not objRS.EOF				
				Set objContentField = new ContentFieldClass
				strID = objRS("id")
				objContentField.setID(strID)
				objContentField.setDescription(objRS("description"))
				strIdGroup = objRS("id_group")
				objContentField.setIdGroup(strIdGroup)
				
				Set objGroup = new ContentFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objContentField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objContentField.setOrder(objRS("order"))	
				objContentField.setTypeField(objRS("type"))
				objContentField.setTypeContent(objRS("type_content"))
				objContentField.setMaxLenght(objRS("max_lenght"))		
				objContentField.setRequired(objRS("required"))	
				objContentField.setEnabled(objRS("enabled"))
				objContentField.setEditable(objRS("editable"))	
				objContentField.setidContent(objRS("id_news"))
				objContentField.setSelValue(objRS("value"))
				objDict.add strID, objContentField
				objRS.moveNext()
			loop
			
			Set objContentField = nothing							
			Set getListContentField4ContentActiveByType = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListContentField4ContentActiveByDesc(idContent, descList)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListContentField4ContentActiveByDesc = null		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder, content_fields_match.id_news, content_fields_match.value FROM content_fields"
		strSQL = strSQL & " LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id"
		strSQL = strSQL & " LEFT JOIN content_fields_match ON content_fields.id = content_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND content_fields_match.id_news=?"	
		if not(isNull(descList)) then
			strSQL = strSQL & " AND ("	
			spitMatchValues = Split(descList,",")
			for j=0 to Ubound(spitMatchValues)
				strSQL = strSQL & " content_fields.description=?"
				if(j<Ubound(spitMatchValues))then
					strSQL = strSQL & " OR "
				end if
			next
			strSQL = strSQL & ")"
		end if		
		strSQL = strSQL & " ORDER BY gorder, content_fields.order;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
		if not(isNull(descList)) then	
			spitMatchValues = Split(descList,",")
			for j=0 to Ubound(spitMatchValues)
				objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,Trim(spitMatchValues(j)))
			next
		end if	
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objContentField
			do while not objRS.EOF				
				Set objContentField = new ContentFieldClass
				strID = objRS("id")
				strDesc = objRS("description")
				objContentField.setID(strID)
				objContentField.setDescription(strDesc)
				strIdGroup = objRS("id_group")
				objContentField.setIdGroup(strIdGroup)
				
				Set objGroup = new ContentFieldGroupClass
				objGroup.setID(strIdGroup)
				objGroup.setDescription(objRS("gdesc"))
				objGroup.setOrder(objRS("gorder"))		
				objContentField.setObjGroup(objGroup)		
				Set objGroup = nothing
				
				objContentField.setOrder(objRS("order"))	
				objContentField.setTypeField(objRS("type"))
				objContentField.setTypeContent(objRS("type_content"))
				objContentField.setMaxLenght(objRS("max_lenght"))		
				objContentField.setRequired(objRS("required"))	
				objContentField.setEnabled(objRS("enabled"))
				objContentField.setEditable(objRS("editable"))	
				objContentField.setidContent(objRS("id_news"))
				objContentField.setSelValue(objRS("value"))
				objDict.add strDesc, objContentField
				objRS.moveNext()
			loop
			
			Set objContentField = nothing							
			Set getListContentField4ContentActiveByDesc = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function hasListContentField4ContentActive(idContent)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		hasListContentField4ContentActive = 0		
		strSQL = "SELECT count(*) as id FROM content_fields"
		strSQL = strSQL & " LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id"
		strSQL = strSQL & " LEFT JOIN content_fields_match ON content_fields.id = content_fields_match.id_field"
		strSQL = strSQL & " WHERE enabled=1"
		strSQL = strSQL & " AND content_fields_match.id_news=?"	
		strSQL = strSQL & " AND (content_fields.type not in(1,2,7,8,9) OR (content_fields.type in(1,2,8,9) AND content_fields.editable=1));"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			hasListContentField4ContentActive = objRS("id")			
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListContentFieldValuesByDesc(strDesc, targetLang, sorting)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListContentFieldValuesByDesc = null		
		strSQL = "SELECT content_fields.id, content_fields_match.value  FROM content_fields"
		strSQL = strSQL & " LEFT JOIN content_fields_match ON content_fields.id = content_fields_match.id_field"
		strSQL = strSQL & " LEFT JOIN target_x_news ON content_fields_match.id_news = target_x_news.id_news"
		strSQL = strSQL & " WHERE enabled=1 AND description=? AND target_x_news.id_target=?"		
		strSQL = strSQL & " GROUP BY value"
		if not(isNull(sorting)) AND sorting<>"" then
			strSQL = strSQL & " ORDER BY CAST(value as "&sorting&");"		
		else
			strSQL = strSQL & " ORDER BY value ASC;"		
		end if
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,targetLang)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			do while not objRS.EOF				
				strID = objRS("id")
				strValue = objRS("value")
				objDict.add strValue, strID
				objRS.moveNext()
			loop
									
			Set getListContentFieldValuesByDesc = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findContentFieldById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findContentFieldById = null		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder FROM content_fields LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id WHERE content_fields.id=?;"
		
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
			Set objContentField = new ContentFieldClass
			strID = objRS("id")
			objContentField.setID(strID)
			objContentField.setDescription(objRS("description"))
			strIdGroup = objRS("id_group")
			objContentField.setIdGroup(strIdGroup)
			
			Set objGroup = new ContentFieldGroupClass
			objGroup.setID(strIdGroup)
			objGroup.setDescription(objRS("gdesc"))
			objGroup.setOrder(objRS("gorder"))
			
			objContentField.setObjGroup(objGroup)
			objContentField.setOrder(objRS("order"))	
			objContentField.setTypeField(objRS("type"))
			objContentField.setTypeContent(objRS("type_content"))
			objContentField.setMaxLenght(objRS("max_lenght"))		
			objContentField.setRequired(objRS("required"))	
			objContentField.setEnabled(objRS("enabled"))
			objContentField.setEditable(objRS("editable"))		
			Set findContentFieldById = objContentField
			Set objContentField = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findContentFieldByDesc(strDesc)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findContentFieldByDesc = null		
		strSQL = "SELECT content_fields.*, content_fields_group.description as gdesc, content_fields_group.order as gorder FROM content_fields LEFT JOIN content_fields_group ON content_fields.id_group=content_fields_group.id WHERE content_fields.description=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strDesc)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objContentField = new ContentFieldClass
			strID = objRS("id")
			objContentField.setID(strID)
			objContentField.setDescription(objRS("description"))
			strIdGroup = objRS("id_group")
			objContentField.setIdGroup(strIdGroup)
			
			Set objGroup = new ContentFieldGroupClass
			objGroup.setID(strIdGroup)
			objGroup.setDescription(objRS("gdesc"))
			objGroup.setOrder(objRS("gorder"))
			
			objContentField.setObjGroup(objGroup)
			objContentField.setOrder(objRS("order"))	
			objContentField.setTypeField(objRS("type"))
			objContentField.setTypeContent(objRS("type_content"))
			objContentField.setMaxLenght(objRS("max_lenght"))		
			objContentField.setRequired(objRS("required"))	
			objContentField.setEnabled(objRS("enabled"))
			objContentField.setEditable(objRS("editable"))		
			Set findContentFieldByDesc = objContentField
			Set objContentField = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
			
	Public Function insertContentField(description, idGroup, order, typeField, typeContent, required, enabled, editable, maxLenght, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		insertContentField = -1
	
		strSQL = "INSERT INTO content_fields(description, id_group, `type`, type_content, `order`, max_lenght, required, enabled, editable) VALUES("
		strSQL = strSQL & "?,"

		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if

		strSQL = strSQL & "?,?,?,"

		if(isNull(maxLenght) OR maxLenght = "") then
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
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		if not isNull(idGroup) AND not(idGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,editable)
		objCommand.Execute
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(content_fields.id) as id FROM content_fields")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertContentField = objRS("id")	
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
		
	Public Sub modifyContentField(id, description, idGroup, order, typeField, typeContent, required, enabled, editable, maxLenght, objConn)
		on error resume next
		Dim objDB, strSQL, objRS

		strSQL = "UPDATE content_fields SET "
		strSQL = strSQL & "description=?,"
		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "id_group=NULL,"
		else
			strSQL = strSQL & "id_group=?,"			
		end if
		strSQL = strSQL & "`type`=?,"
		strSQL = strSQL & "`type_content`=?,"
		strSQL = strSQL & "`order`=?,"	
		if(isNull(maxLenght) OR maxLenght = "") then
			strSQL = strSQL & "max_lenght=NULL,"
		else
			strSQL = strSQL & "max_lenght=?,"			
		end if
		strSQL = strSQL & "required=?,"		
		strSQL = strSQL & "enabled=?,"
		strSQL = strSQL & "editable=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		if not isNull(idGroup) AND not(idGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,editable)
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
			
	Public Function insertContentFieldNoTransaction(description, idGroup, order, typeField, typeContent, required, enabled, editable, maxLenght)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		insertContentFieldNoTransaction = -1
		
		strSQL = "INSERT INTO content_fields(description, id_group, `type`, type_content, `order`, max_lenght, required, enabled, editable) VALUES("
		strSQL = strSQL & "?,"

		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if

		strSQL = strSQL & "?,?,?,"

		if(isNull(maxLenght) OR maxLenght = "") then
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
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		if not isNull(idGroup) AND not(idGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,editable)
		objCommand.Execute()		
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(content_fields.id) as id FROM content_fields")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertContentFieldNoTransaction = objRS("id")	
		end if		
		Set objRS = Nothing		
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyContentFieldNoTransaction(id, description, idGroup, order, typeField, typeContent, required, enabled, editable, maxLenght)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		strSQL = "UPDATE content_fields SET "
		strSQL = strSQL & "description=?,"
		if(isNull(idGroup) OR idGroup = "") then
			strSQL = strSQL & "id_group=NULL,"
		else
			strSQL = strSQL & "id_group=?,"			
		end if
		strSQL = strSQL & "`type`=?,"
		strSQL = strSQL & "`type_content`=?,"
		strSQL = strSQL & "`order`=?,"	
		if(isNull(maxLenght) OR maxLenght = "") then
			strSQL = strSQL & "max_lenght=NULL,"
		else
			strSQL = strSQL & "max_lenght=?,"			
		end if
		strSQL = strSQL & "required=?,"	
		strSQL = strSQL & "enabled=?,"
		strSQL = strSQL & "editable=?"
		strSQL = strSQL & " WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,description)
		if not isNull(idGroup) AND not(idGroup = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,typeContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,order)
		if not isNull(maxLenght) AND not(maxLenght = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,maxLenght)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,enabled)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,editable)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteContentField(id)
		on error resume next
		Dim objDB, strSQL, strSQL2, objRS, objConn		
		strSQL = "DELETE FROM content_fields WHERE id=?;" 
		strSQL2 = "DELETE FROM content_fields_match WHERE id_field=?;"
		strSQL3 = "DELETE FROM content_fields_values WHERE id_field=?;"

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
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQL2
		objCommand3.CommandText = strSQL3
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)
		
		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand2.Execute()
			objCommand3.Execute()
		end if	
		objCommand.Execute()	

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

	
	'************************* GESTIONE PRODUCT FIELDS VALUES *******************************

	Public Function getListContentFieldValues(idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListContentFieldValues = null		
		strSQL = "SELECT * FROM content_fields_values "
		strSQL = strSQL & " WHERE id_field=?"		
		strSQL = strSQL & " ORDER BY `order`;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objContentField
			do while not objRS.EOF				
				strID = objRS("id_field")
				strValue = objRS("value")		
				objDict.add strValue, strID
				objRS.moveNext()
			loop						
			Set getListContentFieldValues = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Sub insertContentFieldValue(idField, strValue, iOrder, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO content_fields_values(id_field, `value`,`order`) VALUES("
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,iOrder)
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
		
	Public Sub modifyContentFieldValue(idField, strValue, iOrder, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE content_fields_values SET "
		strSQL = strSQL & "`value`=?,"
		strSQL = strSQL & "`order`=?"
		strSQL = strSQL & " WHERE id_field=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,iOrder)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
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
		
	Public Sub deleteContentFieldValue(idField, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM content_fields_values WHERE id_field=? AND `value`=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
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
		
	Public Sub deleteContentFieldValueNoTransaction(idField, strValue)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM content_fields_values WHERE id_field=? AND `value`=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub			
		
	Public Sub deleteContentFieldValueByField(idField, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM content_fields_values WHERE id_field=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
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
		
	Public Sub deleteContentFieldValueByFieldNoTransaction(idField)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM content_fields_values WHERE id_field=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub


	'************************* GESTIONE LISTA TYPE *******************************
		
	Public Function getListaTypeField()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaTypeField = null		
		strSQL = "SELECT * FROM content_fields_type;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("description")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaTypeField = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findTypeFieldById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findTypeFieldById = null		
		strSQL = "SELECT * FROM content_fields_type WHERE id=?;"
		
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
			findTypeFieldById = objRS("description")
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	

	'************************* GESTIONE LISTA TYPE CONTENT *******************************
		
	Public Function getListaTypeContent()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaTypeContent = null		
		strSQL = "SELECT * FROM content_fields_type_content;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("description")		
				objDict.add strID, strDesc
				objRS.moveNext()
			loop
							
			Set getListaTypeContent = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findTypeContentById(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findTypeContentById = null		
		strSQL = "SELECT * FROM content_fields_type_content WHERE id=?;"
		
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
			findTypeContentById = objRS("description")
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	'************************* GESTIONE FIELD MATCH *******************************
		
	Public Function findFieldMatch(idField, idContent)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldMatch = null		
		strSQL = "SELECT * FROM content_fields_match WHERE id_field=? AND id_news=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			strID = objRS("id")
			strVal = objRS("value")		
			objDict.add "id", strID	
			objDict.add "value", strVal
			Set findFieldMatch = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findFieldMatchValue(idField, idContent)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		findFieldMatchValue = null		
		strSQL = "SELECT * FROM content_fields_match WHERE id_field=? AND id_news=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,idContent)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then
			strVal = objRS("value")	
			findFieldMatchValue = strVal		
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertFieldMatch(idField, idContent, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO content_fields_match(id_field, id_news, `value`) VALUES("
		strSQL = strSQL & "?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idContent)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
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
		
	Public Sub modifyFieldMatch(idField, idContent, strValue, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE content_fields_match SET "
		strSQL = strSQL & "`value`=?"
		strSQL = strSQL & " WHERE id_field=? AND id_news=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strValue)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idField)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idContent)
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
		
	Public Sub deleteFieldMatch(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM content_fields_match WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
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
		
	Public Sub deleteFieldMatchNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM content_fields_match WHERE id=?;"

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
		
	Public Sub deleteFieldMatchByContent(idContent, objConn)
		on error resume next
		Dim objDB, strSQL, objRS		
		strSQL = "DELETE FROM content_fields_match WHERE id_news=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idContent)
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
		
	Public Sub deleteFieldMatchByContentNoTransaction(idContent)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM content_fields_match WHERE id_news=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idContent)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	

	'************************* GESTIONE FIELD RENDERING *******************************
	
	Public Function renderContentFieldHTML(contentField,cssClass, customFieldPrefix, idContent, defaultMatchValue, translator, isClient, isEditable)	
		Dim fieldMatchValue, spitValues, keyPress, maxLenght, style
		
		fieldMatchValue = defaultMatchValue
		
		on error resume next			
		keyPress = ""
		select Case contentField.getTypeContent()
		Case 2		
			keyPress = " onkeypress=""javascript:return isInteger(event);"""
		Case 3		
			keyPress = " onkeypress=""javascript:return isDouble(event);"""
		Case Else
		End Select		
		
		maxLenght = ""		
		if not(contentField.getMaxLenght()="") AND (contentField.getMaxLenght()>0) then
			maxLenght = " maxlength="""&contentField.getMaxLenght()&""""
		end if
		
		style = ""		
		if not(cssClass="") then
			style = " class="""&cssClass&""""
		end if

		if isNull(customFieldPrefix) then
			customFieldPrefix = ""
		end if
		
		renderContentFieldHTML = ""		

		select Case contentField.getTypeField()
		Case 1
			if(Trim(fieldMatchValue)<>"")then fieldMatchValue=Server.HTMLEncode(fieldMatchValue) end if
			
			if(isClient) then
				if(isEditable)then 
					renderContentFieldHTML = "<input type=""text"" name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" value="""&fieldMatchValue&""" "&style&" "&keyPress&maxLenght&"/>"
	
					if(CInt(contentField.getTypeContent())=4) then
						renderContentFieldHTML = renderContentFieldHTML & "<script>"
						renderContentFieldHTML = renderContentFieldHTML & "$(function() {"
							renderContentFieldHTML = renderContentFieldHTML & "$('#"&getFieldPrefix()&customFieldPrefix&contentField.getID()&"').datepicker({"
								renderContentFieldHTML = renderContentFieldHTML & "dateFormat: 'dd/mm/yy',"
								renderContentFieldHTML = renderContentFieldHTML & "changeMonth: true,"
								renderContentFieldHTML = renderContentFieldHTML & "changeYear: true"
								'renderContentFieldHTML = renderContentFieldHTML & ",yearRange: '1900:"&DatePart("yyyy",Date())&"'" 
							renderContentFieldHTML = renderContentFieldHTML & "});"
						renderContentFieldHTML = renderContentFieldHTML & "});"
						renderContentFieldHTML = renderContentFieldHTML & "</script>"
					end if					
				else
					if not(translator.getTranslated(fieldMatchValue)="") then fieldMatchValue=translator.getTranslated(fieldMatchValue) end if 
					renderContentFieldHTML = fieldMatchValue					
				end if
			else
				if(isEditable)then
					renderContentFieldHTML = fieldMatchValue
				else		
					renderContentFieldHTML = "<input type=""text"" name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" value="""&fieldMatchValue&""" "&style&" "&keyPress&maxLenght&"/>"
	
					if(CInt(contentField.getTypeContent())=4) then
						renderContentFieldHTML = renderContentFieldHTML & "<script>"
						renderContentFieldHTML = renderContentFieldHTML & "$(function() {"
							renderContentFieldHTML = renderContentFieldHTML & "$('#"&getFieldPrefix()&customFieldPrefix&contentField.getID()&"').datepicker({"
							renderContentFieldHTML = renderContentFieldHTML & "dateFormat: 'dd/mm/yy',"
							renderContentFieldHTML = renderContentFieldHTML & "changeMonth: true,"
							renderContentFieldHTML = renderContentFieldHTML & "changeYear: true"
							renderContentFieldHTML = renderContentFieldHTML & "});"
						renderContentFieldHTML = renderContentFieldHTML & "});"
						renderContentFieldHTML = renderContentFieldHTML & "</script>"
					end if			
				end if
			end if
		Case 2
			if(Trim(fieldMatchValue)<>"")then fieldMatchValue=Server.HTMLEncode(fieldMatchValue) end if
			
			if(isClient) then
				if(isEditable)then
					renderContentFieldHTML = "<textarea name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
				else
					if not(translator.getTranslated(fieldMatchValue)="") then fieldMatchValue=translator.getTranslated(fieldMatchValue) end if 
					renderContentFieldHTML = fieldMatchValue
				end if
			else
				if(isEditable)then
					renderContentFieldHTML = fieldMatchValue
				else
					renderContentFieldHTML = "<textarea name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
				end if
			end if
		Case 3
			Dim key, objCountry 
				
			renderContentFieldHTML = "<select name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" "&style&">"
			renderContentFieldHTML = renderContentFieldHTML & "<option value=""""></option>"			
			
			if(CInt(contentField.getTypeContent())=5) then
				On Error Resume Next
				Set objCountry = New CountryClass				
				Set specialFieldValue = objCountry.findCountryListOnly("1,3")
				if (strComp(typename(specialFieldValue), "Dictionary") = 0)then
					for each x in specialFieldValue
						key =  specialFieldValue(x).getCountryCode()&"_"&specialFieldValue(x).getID()
						selected = ""
						label = specialFieldValue(x).getCountryDescription()
						if not(translator.getTranslated("portal.commons.select.option.country."&specialFieldValue(x).getCountryCode())="") then label=translator.getTranslated("portal.commons.select.option.country."&specialFieldValue(x).getCountryCode()) end if
						if (strComp(key, fieldMatchValue, 1) = 0) then selected=" selected" end if
						renderContentFieldHTML = renderContentFieldHTML & "<option value="""&key&""" "&selected&">"&label&"</option>"     
					next
				end if
				Set specialFieldValue = nothing					
				Set objCountry = nothing
				if(Err.number <> 0)then
				end if
			elseif(CInt(contentField.getTypeContent())=6)then
				On Error Resume Next
				Set objCountry = New CountryClass
				Set specialFieldValue = objCountry.findStateRegionListOnly("1,3")
				if (strComp(typename(specialFieldValue), "Dictionary") = 0)then
					for each x in specialFieldValue
						key =  specialFieldValue(x).getStateRegionCode()&"_"&specialFieldValue(x).getID()
						selected = ""
						label = specialFieldValue(x).getStateRegionDescription()
						if not(translator.getTranslated("portal.commons.select.option.country."&specialFieldValue(x).getStateRegionCode())="") then label=translator.getTranslated("portal.commons.select.option.country."&specialFieldValue(x).getStateRegionCode()) end if
						if (strComp(key, fieldMatchValue, 1) = 0) then selected=" selected" end if
						renderContentFieldHTML = renderContentFieldHTML & "<option value="""&key&""" "&selected&">"&label&"</option>"     
					next
				end if
				Set specialFieldValue = nothing
				Set objCountry = nothing
				if(Err.number <> 0)then
				end if			
			else
				On Error Resume Next
				spitValues = getListContentFieldValues(contentField.getID()).Keys
				for each x in spitValues			
					selected = ""
					if (strComp(Trim(x), fieldMatchValue, 1) = 0) then selected=" selected" end if
					label= Trim(x)
					if not(translator.getTranslated("portal.commons.content_field.label."&label)="") then label=translator.getTranslated("portal.commons.content_field.label."&label) end if
					renderContentFieldHTML = renderContentFieldHTML & "<OPTION VALUE="""&x&""""&selected&">"&label&"</OPTION>"
				next
				if(Err.number <> 0)then
				end if			
			end if
			
			renderContentFieldHTML = renderContentFieldHTML & "</select>"
		Case 4
			renderContentFieldHTML ="<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&customFieldPrefix&contentField.getID()&""">"
			
			renderContentFieldHTML = renderContentFieldHTML & "<select name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" multiple size="""&contentField.getMaxLenght()&""" "&style&">"
			renderContentFieldHTML = renderContentFieldHTML & "<option value=""""></option>"			
			
			On Error Resume Next
			spitValues = getListContentFieldValues(contentField.getID()).Keys
			for each x in spitValues
				selected = ""
				'if (strComp(Trim(x), fieldMatchValue, 1) = 0) then selected=" selected" end if
				
				if not(fieldMatchValue = "") then
					spitMatchValues = Split(fieldMatchValue,",")
					for j=0 to Ubound(spitMatchValues)
						if(strComp(Trim(spitMatchValues(j)), Trim(x), 1) = 0) then
							selected=" selected"
							exit for
						end if
					next
				end if				
				
				label= Trim(x)
				if not(translator.getTranslated("portal.commons.content_field.label."&label)="") then label=translator.getTranslated("portal.commons.content_field.label."&label) end if
				renderContentFieldHTML = renderContentFieldHTML & "<OPTION VALUE="""&x&""""&selected&">"&label&"</OPTION>"
			next
			if(Err.number <> 0)then
			end if
			
			renderContentFieldHTML = renderContentFieldHTML & "</select>"
		Case 5			
			on error Resume Next
			spitValues = getListContentFieldValues(contentField.getID()).Keys
			K=1		
	
			renderContentFieldHTML =renderContentFieldHTML & "<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&customFieldPrefix&contentField.getID()&""">"

			for each y in spitValues
				checked = ""
				if not(fieldMatchValue = "") then
					spitMatchValues = Split(fieldMatchValue,",")
					for j=0 to Ubound(spitMatchValues)
						if(strComp(Trim(spitMatchValues(j)), Trim(y), 1) = 0) then
							checked=" checked='checked'"
							exit for
						end if
					next
				end if
				newLine = ""
				if((k Mod 4) = 0) then newLine="<br/>" end if
				label= Trim(y)
				if not(translator.getTranslated("portal.commons.content_field.label."&label)="") then label=translator.getTranslated("portal.commons.content_field.label."&label) end if
				renderContentFieldHTML =renderContentFieldHTML & label&"&nbsp;<input type=""checkbox"" "&style&" value="""&y&""" name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" "&checked&"/>"&newLine
				k = k+1
			next
			
			if Err.number <> 0 then
				'response.write(Err.description)
			end if
		Case 6			
			on error Resume Next
			spitValues = getListContentFieldValues(contentField.getID()).Keys
			K=1
			
			renderContentFieldHTML =renderContentFieldHTML & "<input type=""hidden"" value="""" name=""hidden_"&getFieldPrefix()&customFieldPrefix&contentField.getID()&""">"

			for each y in spitValues
				checked = ""
				if not(fieldMatchValue = "") then
					if (strComp(fieldMatchValue, Trim(y), 1) = 0) then checked=" checked='checked'" end if
				end if
				newLine = ""
				if((k Mod 4) = 0) then newLine="<br/>" end if
				label= Trim(y)
				if not(translator.getTranslated("portal.commons.content_field.label."&label)="") then label=translator.getTranslated("portal.commons.content_field.label."&label) end if
				renderContentFieldHTML =renderContentFieldHTML & label&"&nbsp;<input type=""radio"" "&style&" value="""&y&""" name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" "&checked&"/>"&newLine
				k = k+1
			next
			
			if Err.number <> 0 then
				'response.write(Err.description)
			end if
		Case 7
			if(Trim(fieldMatchValue)<>"")then fieldMatchValue=Server.HTMLEncode(fieldMatchValue) end if
			
			renderContentFieldHTML = "<input type=""hidden"" value="""&fieldMatchValue&""" name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&"""/>"
		Case 8
			renderContentFieldHTML = "<input type=""file"" name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" "&style&" />"
		Case 9
			if(Trim(fieldMatchValue)<>"")then fieldMatchValue=Server.HTMLEncode(fieldMatchValue) end if
			
			if(isClient) then
				if(isEditable)then
					renderContentFieldHTML = "<script type=""text/javascript"">"
					renderContentFieldHTML = renderContentFieldHTML & "$.cleditor.defaultOptions.width = 280;"
					renderContentFieldHTML = renderContentFieldHTML & "$.cleditor.defaultOptions.height = 200;"
					renderContentFieldHTML = renderContentFieldHTML & "$.cleditor.defaultOptions.controls = ""bold italic underline strikethrough subscript superscript | font size style | color highlight removeformat | bullets numbering | alignleft center alignright justify | rule | cut copy paste | image"";"		
					renderContentFieldHTML = renderContentFieldHTML & "$(document).ready(function() {$(""#"&getFieldPrefix()&customFieldPrefix&contentField.getID()&""").cleditor();});"
					renderContentFieldHTML = renderContentFieldHTML & "</script>"
					renderContentFieldHTML = renderContentFieldHTML & "<textarea name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
				else
					if not(translator.getTranslated(fieldMatchValue)="") then fieldMatchValue=translator.getTranslated(fieldMatchValue) end if 
					renderContentFieldHTML = fieldMatchValue
				end if
			else
				if(isEditable)then
					renderContentFieldHTML = fieldMatchValue
				else
					renderContentFieldHTML = "<script type=""text/javascript"">"
					renderContentFieldHTML = renderContentFieldHTML & "$.cleditor.defaultOptions.width = 280;"
					renderContentFieldHTML = renderContentFieldHTML & "$.cleditor.defaultOptions.height = 200;"
					renderContentFieldHTML = renderContentFieldHTML & "$.cleditor.defaultOptions.controls = ""bold italic underline strikethrough subscript superscript | font size style | color highlight removeformat | bullets numbering | alignleft center alignright justify | rule | cut copy paste | image"";"		
					renderContentFieldHTML = renderContentFieldHTML & "$(document).ready(function() {$(""#"&getFieldPrefix()&customFieldPrefix&contentField.getID()&""").cleditor();});"
					renderContentFieldHTML = renderContentFieldHTML & "</script>"
					renderContentFieldHTML = renderContentFieldHTML & "<textarea name="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" id="""&getFieldPrefix()&customFieldPrefix&contentField.getID()&""" "&style&" >"&fieldMatchValue&"</textarea>"
				end if
			end if				
		Case Else
		End Select	
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function
	
	Public Function renderContentFieldJS(contentField, formName, customFieldPrefix, translator,defaultMatchValue, frontOrBack)	
		on error resume next

		if isNull(customFieldPrefix) then
			customFieldPrefix = ""
		end if
		
		renderContentFieldJS = ""	
		

		select Case contentField.getTypeField()
		Case 1,2
			renderContentFieldJS = "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"){"
			renderContentFieldJS = renderContentFieldJS & "var "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_default_match_values = """&defaultMatchValue&""";"
			renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value.toLowerCase() == "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_default_match_values.toLowerCase()){"
				renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value = """";"
			renderContentFieldJS = renderContentFieldJS &"}"				
		
			if(CInt(contentField.getRequired())=1)then
				'se backoffice, verifico se è stata selezionata la checkbox del field e solo in quel caso attivo il controllo required
				if(frontOrBack="1")then
					'caso frontend, do nothing
				elseif(frontOrBack="2")then
					'caso backend
					renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active != null){"
						renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active.length == null){"
							renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active.checked){"
								renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value == """"){"
									renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.insert_"&contentField.getDescription())&""");"
									renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".focus();"
									renderContentFieldJS = renderContentFieldJS &"return false;"
								renderContentFieldJS = renderContentFieldJS &"}"
							renderContentFieldJS = renderContentFieldJS & "}"
						renderContentFieldJS = renderContentFieldJS & "}else{"
							renderContentFieldJS = renderContentFieldJS & "for(k=0; k<document."&formName&".content_field_active.length; k++){"
								renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active[k].checked && document."&formName&".content_field_active[k].value=="""&contentField.getID()&"-"&contentField.getTypeField()&"""){"		
									renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value == """"){"
										renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.insert_"&contentField.getDescription())&""");"
										renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".focus();"
										renderContentFieldJS = renderContentFieldJS &"return false;"
									renderContentFieldJS = renderContentFieldJS &"}"
								renderContentFieldJS = renderContentFieldJS & "}"
							renderContentFieldJS = renderContentFieldJS & "}"
						renderContentFieldJS = renderContentFieldJS & "}"
					renderContentFieldJS = renderContentFieldJS & "}"					
				end if
			end if

			if(CInt(contentField.getTypeContent())=2) then
				renderContentFieldJS = renderContentFieldJS &"if(isNaN(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value)){"
					renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.isnan_value")&""");"
					renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".focus();"
					renderContentFieldJS = renderContentFieldJS &"return false;"	
				renderContentFieldJS = renderContentFieldJS &"}"			
			end if

			if(CInt(contentField.getTypeContent())=3) then
				renderContentFieldJS = renderContentFieldJS &"if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value.length > 0 && (!checkDoubleFormatExt(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value) || document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value.indexOf(""."")!=-1)){"
					renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.isnan_value")&""");"
					renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".focus();"
					renderContentFieldJS = renderContentFieldJS &"return false;"
				renderContentFieldJS = renderContentFieldJS &"}"		
			end if
			renderContentFieldJS = renderContentFieldJS &"}"

		Case 3
			if(CInt(contentField.getRequired())=1)then	
				'se backoffice, verifico se è stata selezionata la checkbox del field e solo in quel caso attivo il controllo required
				if(frontOrBack="1")then
					'caso frontend, do nothing
				elseif(frontOrBack="2")then
					renderContentFieldJS = "if(document."&formName&".content_field_active != null){"
						renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active.length == null){"
							renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active.checked){"			
								renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"){"
									renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options[document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".selectedIndex].value == """"){"
										renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.insert_"&contentField.getDescription())&""");"
										renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".focus();"
										renderContentFieldJS = renderContentFieldJS &"return false;"
									renderContentFieldJS = renderContentFieldJS &"}"
								renderContentFieldJS = renderContentFieldJS &"}"
							renderContentFieldJS = renderContentFieldJS & "}"
						renderContentFieldJS = renderContentFieldJS & "}else{"
							renderContentFieldJS = renderContentFieldJS & "for(k=0; k<document."&formName&".content_field_active.length; k++){"
								renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active[k].checked && document."&formName&".content_field_active[k].value=="""&contentField.getID()&"-"&contentField.getTypeField()&"""){"	
									renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"){"
										renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options[document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".selectedIndex].value == """"){"
											renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.insert_"&contentField.getDescription())&""");"
											renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".focus();"
											renderContentFieldJS = renderContentFieldJS &"return false;"
										renderContentFieldJS = renderContentFieldJS &"}else{break;}"
									renderContentFieldJS = renderContentFieldJS &"}"
								renderContentFieldJS = renderContentFieldJS & "}"
							renderContentFieldJS = renderContentFieldJS & "}"						
						renderContentFieldJS = renderContentFieldJS & "}"
					renderContentFieldJS = renderContentFieldJS & "}"						
				end if								
			end if	

		Case 4	
			renderContentFieldJS = ""
			if(CInt(contentField.getRequired())=1)then	
				'se backoffice, verifico se è stata selezionata la checkbox del field e solo in quel caso attivo il controllo required
				if(frontOrBack="1")then
					'caso frontend, do nothing
				elseif(frontOrBack="2")then
					renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active != null){"
						renderContentFieldJS = renderContentFieldJS & "var "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_hasselection = false;"
						renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active.length == null){"
							renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active.checked){"			
								renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"){"
									renderContentFieldJS = renderContentFieldJS & "for(k=0; k<document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options.length; k++){"
										renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options[k].selected && document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options[k].value != """"){"							
											renderContentFieldJS = renderContentFieldJS &getFieldPrefix()&customFieldPrefix&contentField.getID()&"_hasselection = true;"
											renderContentFieldJS = renderContentFieldJS &"break;"
										renderContentFieldJS = renderContentFieldJS & "}"					
									renderContentFieldJS = renderContentFieldJS & "}"
									renderContentFieldJS = renderContentFieldJS & "if(!"&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_hasselection){"
										renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.insert_"&contentField.getDescription())&""");"
										renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".focus();"
										renderContentFieldJS = renderContentFieldJS &"return false;"
									renderContentFieldJS = renderContentFieldJS &"}"									
								renderContentFieldJS = renderContentFieldJS &"}"
							renderContentFieldJS = renderContentFieldJS & "}"
						renderContentFieldJS = renderContentFieldJS & "}else{"
							renderContentFieldJS = renderContentFieldJS & "for(k=0; k<document."&formName&".content_field_active.length; k++){"
								renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active[k].checked && document."&formName&".content_field_active[k].value=="""&contentField.getID()&"-"&contentField.getTypeField()&"""){"	
									renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"){"
										renderContentFieldJS = renderContentFieldJS & "for(k=0; k<document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options.length; k++){"
											renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options[k].selected && document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options[k].value != """"){"							
												renderContentFieldJS = renderContentFieldJS &getFieldPrefix()&customFieldPrefix&contentField.getID()&"_hasselection = true;"
												renderContentFieldJS = renderContentFieldJS &"break;"
											renderContentFieldJS = renderContentFieldJS & "}"					
										renderContentFieldJS = renderContentFieldJS & "}"										
										renderContentFieldJS = renderContentFieldJS & "if(!"&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_hasselection){"
											renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.insert_"&contentField.getDescription())&""");"
											renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".focus();"
											renderContentFieldJS = renderContentFieldJS &"return false;"
										renderContentFieldJS = renderContentFieldJS &"}else{break;}"										
									renderContentFieldJS = renderContentFieldJS &"}"
								renderContentFieldJS = renderContentFieldJS & "}"
							renderContentFieldJS = renderContentFieldJS & "}"						
						renderContentFieldJS = renderContentFieldJS & "}"
					renderContentFieldJS = renderContentFieldJS & "}"						
				end if								
			end if
			
			if(frontOrBack="1")then
				'caso frontend, do nothing
			elseif(frontOrBack="2")then
				renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"){"
					renderContentFieldJS = renderContentFieldJS & "var "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values = """";"
					renderContentFieldJS = renderContentFieldJS & "for(k=0; k<document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options.length; k++){"
						renderContentFieldJS = renderContentFieldJS & "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options[k].selected){"							
							renderContentFieldJS = renderContentFieldJS &getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values = "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values + document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".options[k].value + "","";"
						renderContentFieldJS = renderContentFieldJS & "}"					
					renderContentFieldJS = renderContentFieldJS & "}"
					renderContentFieldJS = renderContentFieldJS &getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values = "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values.substring(0, "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values.lastIndexOf(','));"
					renderContentFieldJS = renderContentFieldJS &"document."&formName&".hidden_"&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value = "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values;"	
				renderContentFieldJS = renderContentFieldJS & "}"				
			end if			
		Case 5,6
			renderContentFieldJS = "if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"){"
			renderContentFieldJS = renderContentFieldJS & "var "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values = """";"
			renderContentFieldJS = renderContentFieldJS &"if (document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"){"
				renderContentFieldJS = renderContentFieldJS &"if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&" != null){"				
					renderContentFieldJS = renderContentFieldJS &"if(document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".length == null){"
						renderContentFieldJS = renderContentFieldJS &"if (document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".checked){"
							renderContentFieldJS = renderContentFieldJS &getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values = "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values + document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value + "","";"
							renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".checked=false;"
						renderContentFieldJS = renderContentFieldJS &"}"						
					renderContentFieldJS = renderContentFieldJS &"}else{"
						renderContentFieldJS = renderContentFieldJS &"for (var i=0; i < document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&".length; i++){"
							renderContentFieldJS = renderContentFieldJS &"if (document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"[i].checked){"
								renderContentFieldJS = renderContentFieldJS &getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values = "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values + document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"[i].value + "","";"
								renderContentFieldJS = renderContentFieldJS &"document."&formName&"."&getFieldPrefix()&customFieldPrefix&contentField.getID()&"[i].checked=false;"
							renderContentFieldJS = renderContentFieldJS &"}"
						renderContentFieldJS = renderContentFieldJS &"}"						
					renderContentFieldJS = renderContentFieldJS &"}"						
					renderContentFieldJS = renderContentFieldJS &getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values = "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values.substring(0, "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values.lastIndexOf(','));"
				renderContentFieldJS = renderContentFieldJS &"}"
			renderContentFieldJS = renderContentFieldJS &"}"
			renderContentFieldJS = renderContentFieldJS &"document."&formName&".hidden_"&getFieldPrefix()&customFieldPrefix&contentField.getID()&".value = "&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values;"
								
			if(CInt(contentField.getRequired())=1)then	
				'se backoffice, verifico se è stata selezionata la checkbox del field e solo in quel caso attivo il controllo required
				if(frontOrBack="1")then
					'caso frontend, do nothing
				elseif(frontOrBack="2")then		
					renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active != null){"
						renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active.length == null){"
							renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active.checked){"			
								renderContentFieldJS = renderContentFieldJS &"if ("&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values ==""""){"
									renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.insert_"&contentField.getDescription())&""");"
									renderContentFieldJS = renderContentFieldJS &"return false;"
								renderContentFieldJS = renderContentFieldJS &"}"
							renderContentFieldJS = renderContentFieldJS & "}"
						renderContentFieldJS = renderContentFieldJS & "}else{"
							renderContentFieldJS = renderContentFieldJS & "for(k=0; k<document."&formName&".content_field_active.length; k++){"
								renderContentFieldJS = renderContentFieldJS & "if(document."&formName&".content_field_active[k].checked && document."&formName&".content_field_active[k].value=="""&contentField.getID()&"-"&contentField.getTypeField()&"""){"	
									renderContentFieldJS = renderContentFieldJS &"if ("&getFieldPrefix()&customFieldPrefix&contentField.getID()&"_values ==""""){"
										renderContentFieldJS = renderContentFieldJS &"alert("""&translator.getTranslated("portal.commons.content_field.js.alert.insert_"&contentField.getDescription())&""");"
										renderContentFieldJS = renderContentFieldJS &"return false;"
									renderContentFieldJS = renderContentFieldJS &"}"
								renderContentFieldJS = renderContentFieldJS & "}"
							renderContentFieldJS = renderContentFieldJS & "}"						
						renderContentFieldJS = renderContentFieldJS & "}"
					renderContentFieldJS = renderContentFieldJS & "}"
				end if						
			end if
			renderContentFieldJS = renderContentFieldJS &"}"
		Case Else
		End Select
 
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Function
	
	Public Function getFieldPrefix()
		getFieldPrefix = "contentfield"
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
End Class
%>