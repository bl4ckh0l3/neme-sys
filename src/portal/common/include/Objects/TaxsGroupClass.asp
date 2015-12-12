<%

Class TaxsGroupClass
	Private id
	Private group_description
	Private country_code
	Private state_region_code
	Private id_tax
	Private exclude_calculation
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub
	
	Public Function getGroupDescription()
		getGroupDescription = group_description
	End Function
	
	Public Sub setGroupDescription(strDesc)
		group_description = strDesc
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
	
	Public Function getTaxID()
		getTaxID = id_tax
	End Function
	
	Public Sub setTaxID(strTaxID)
		id_tax = strTaxID
	End Sub
	
	Public Function isExcludeCalculation()
		isExcludeCalculation = exclude_calculation
	End Function
	
	Public Sub setExcludeCalculation(strExcludeCalculation)
		exclude_calculation = strExcludeCalculation
	End Sub
	
		
	Public Function getListaTaxsGroup(groupDesc)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaTaxsGroup = null		
		strSQL = "SELECT * FROM tax_group"
		if (not(isNull(groupDesc)) AND groupDesc<>"") then strSQL = strSQL & " WHERE description LIKE ?"
		strSQL = strSQL & " ORDER BY description;"
		strSQL = Trim(strSQL)


		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		if (not(isNull(groupDesc)) AND groupDesc<>"") then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&groupDesc&"%") end if
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("description")
						
				Set objTaxGTmp = new TaxsGroupClass
				objTaxGTmp.setID(strID)
				objTaxGTmp.setGroupDescription(strDesc)								
				objDict.add strID, objTaxGTmp
				Set objTaxGTmp = Nothing
				objRS.moveNext()
			loop
							
			Set getListaTaxsGroup = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function getGroupByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getGroupByID = null		
		strSQL = "SELECT * FROM tax_group WHERE id=?;"

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
			strID = objRS("id")
			strDesc = objRS("description")
					
			Set objTaxGTmp = new TaxsGroupClass
			objTaxGTmp.setID(strID)
			objTaxGTmp.setGroupDescription(strDesc)						
			Set getGroupByID = objTaxGTmp		
			Set objTaxGTmp = Nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function findWrapTaxsGroup(groupID)
		on error resume next
		findWrapTaxsGroup = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT tax_group.id, tax_group.description, tax_group_value.country_code, tax_group_value.state_region_code, tax_group_value.id_tassa_applicata, tax_group_value.exclude_calculation  FROM tax_group LEFT JOIN tax_group_value ON tax_group.id=tax_group_value.id_group"
		if (not(isNull(groupID)) AND groupID<>"") then		
			strSQL = strSQL & " WHERE id=?"
		end if
		strSQL = strSQL & " ORDER BY description, country_code, state_region_code"
		strSQL = strSQL & ";"
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		if (not(isNull(groupID)) AND groupID<>"") then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,groupID) end if

		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			refID = ""
			
			do while not objRS.EOF
				strID = objRS("id")
				strDesc = objRS("description")
				strCountryCode = objRS("country_code")
				strStateRegionCode = objRS("state_region_code")
				strTaxID = objRS("id_tassa_applicata")
				bolExcludeCalculation = objRS("exclude_calculation")

				Set objTaxGTmp = new TaxsGroupClass
				objTaxGTmp.setID(strID)
				objTaxGTmp.setGroupDescription(strDesc)
				objTaxGTmp.setCountryCode(strCountryCode)	
				objTaxGTmp.setStateRegionCode(strStateRegionCode)
				objTaxGTmp.setTaxID(strTaxID)
				objTaxGTmp.setExcludeCalculation(bolExcludeCalculation)

				composedId = strID&"|"&Trim(strDesc)
				
				if(Cint(composedId)=cint(refID)) then
					if(objDict.Exists(composedId))then
						objDict(composedId).add objTaxGTmp, ""
					else
						if(isNull(strCountryCode) OR Trim(strCountryCode)="")then
							objDictWrap = null
						else					
							Set objDictWrap = Server.CreateObject("Scripting.Dictionary")
							objDictWrap.add objTaxGTmp, ""
						end if
						objDict.add composedId, objDictWrap							
					end if
				else
					if(isNull(strCountryCode) OR Trim(strCountryCode)="")then
						objDictWrap = null
					else	
						Set objDictWrap = Server.CreateObject("Scripting.Dictionary")
						objDictWrap.add objTaxGTmp, ""
					end if
					objDict.add composedId, objDictWrap						
				end if

				refID = composedId
				Set objTaxGTmp = Nothing
				objRS.moveNext()
			loop
							
			Set findWrapTaxsGroup = objDict			
			Set objDict = nothing
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findTaxsGroupByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findTaxsGroupByID = null
		strSQL = "SELECT tax_group.id, tax_group.description, tax_group_value.country_code, tax_group_value.state_region_code, tax_group_value.id_tassa_applicata, tax_group_value.exclude_calculation FROM tax_group LEFT JOIN tax_group_value ON tax_group.id=tax_group_value.id_group WHERE tax_group.id=?;"		
		
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
			Set objDict = Server.CreateObject("Scripting.Dictionary")			
			do while not objRS.EOF			
				strID = objRS("id_group")
				strDesc = objRS("description")
				strCountryCode = objRS("country_code")
				strStateRegionCode = objRS("state_region_code")
				strTaxID = objRS("id_tassa_applicata")
				bolExcludeCalculation = objRS("exclude_calculation")
						
				Set objTaxGTmp = new TaxsGroupClass
				objTaxGTmp.setID(strID)
				objTaxGTmp.setGroupDescription(strDesc)
				objTaxGTmp.setCountryCode(strCountryCode)	
				objTaxGTmp.setStateRegionCode(strStateRegionCode)
				objTaxGTmp.setTaxID(strTaxID)
				objTaxGTmp.setExcludeCalculation(bolExcludeCalculation)
				objDict.add strID, objTaxGTmp
				Set objTaxGTmp = Nothing
				objRS.moveNext()
			loop
							
			Set findTaxsGroupByID = objDict			
			Set objDict = nothing	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function findTaxsGroupValue(groupID, countryCode, stateRegionCode)
		on error resume next
		findTaxsGroupValue = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT tax_group.id, tax_group.description, tax_group_value.country_code, tax_group_value.state_region_code, tax_group_value.id_tassa_applicata, tax_group_value.exclude_calculation FROM tax_group LEFT JOIN tax_group_value ON tax_group.id=tax_group_value.id_group WHERE tax_group.id=? AND country_code=?"

		if (not(isNull(stateRegionCode)) AND stateRegionCode<>"") then 
			strSQL = strSQL & " AND state_region_code=?"
		else
			strSQL = strSQL & " AND ISNULL(state_region_code)"
		end if

		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,groupID)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode)
		if not isNull(stateRegionCode) AND not(stateRegionCode = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode)
		end if		
		Set objRS = objCommand.Execute()
	
		if not(objRS.EOF) then		
			strID = objRS("id_group")
			strDesc = objRS("description")
			strCountryCode = objRS("country_code")
			strStateRegionCode = objRS("state_region_code")
			strTaxID = objRS("id_tassa_applicata")
			bolExcludeCalculation = objRS("exclude_calculation")
					
			Set objTaxGTmp = new TaxsGroupClass
			objTaxGTmp.setID(strID)
			objTaxGTmp.setGroupDescription(strDesc)
			objTaxGTmp.setCountryCode(strCountryCode)	
			objTaxGTmp.setStateRegionCode(strStateRegionCode)
			objTaxGTmp.setTaxID(strTaxID)
			objTaxGTmp.setExcludeCalculation(bolExcludeCalculation)								
			Set findTaxsGroupValue = objTaxGTmp
			Set objTaxGTmp = nothing			
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing	
		Set objDB = Nothing	
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function findRelatedTax(groupID, countryCode, stateRegionCode)
		on error resume next
		findRelatedTax = null		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT id_tassa_applicata FROM tax_group_value WHERE id_group=? AND country_code=?"

		if (not(isNull(stateRegionCode)) AND stateRegionCode<>"") then 
			strSQL = strSQL & " AND state_region_code=?"
		else
			strSQL = strSQL & " AND ISNULL(state_region_code)"
		end if

		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,groupID)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode)
		if not isNull(stateRegionCode) AND not(stateRegionCode = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode)
		end if		
		Set objRS = objCommand.Execute()
	
		if not(objRS.EOF) then		
			strTaxID = objRS("id_tassa_applicata")									
			findRelatedTax = strTaxID	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing	
		Set objDB = Nothing	
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function isTaxExclusion(groupID, countryCode, stateRegionCode)
		on error resume next
		isTaxExclusion = false		

		Dim objDB, strSQL, objRS, objConn, objDict
		strSQL = "SELECT id_tassa_applicata FROM tax_group_value WHERE id_group=? AND country_code=?"

		if (not(isNull(stateRegionCode)) AND stateRegionCode<>"") then 
			strSQL = strSQL & " AND state_region_code=?"
		else
			strSQL = strSQL & " AND ISNULL(state_region_code)"
		end if
		strSQL = strSQL & " AND exclude_calculation=1"
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,groupID)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode)
		if not isNull(stateRegionCode) AND not(stateRegionCode = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode)
		end if		
		Set objRS = objCommand.Execute()
	
		if not(objRS.EOF) then										
			isTaxExclusion = true	
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing	
		Set objDB = Nothing	
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function getImportoTassa(dblPrezzo, objTax)
		dim importo, iValore		
		iValore = objTax.getValore()
		iValore = CDbl(iValore)
		if(objTax.getTipoValore() = 2) then
			importo = CDbl(dblPrezzo) * (iValore / 100)
		else
			importo = iValore
		end if
		
		getImportoTassa = importo
	End Function
				
	Public Sub insertTaxsGroup(strDesc, objConn)
		on error resume next
		Dim strSQL
		
		strSQL = "INSERT INTO tax_group(description) VALUES("
		strSQL = strSQL & "?);"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strDesc)
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
				
	Public Sub insertTaxsGroupValue(idGroup, countryCode, stateRegionCode, taxID, excludeCalculation, objConn)
		on error resume next
		Dim strSQL
		
		strSQL = "INSERT INTO tax_group_value(id_group, country_code, state_region_code, id_tassa_applicata, exclude_calculation) VALUES("
		strSQL = strSQL & "?,?,"
		if(isNull(stateRegionCode) OR stateRegionCode = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"			
		end if
		if(isNull(taxID) OR taxID = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"			
		end if	
		strSQL = strSQL & ",?);"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode)
		if not isNull(stateRegionCode) AND not(stateRegionCode = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode)
		end if
		if not isNull(taxID) AND not(taxID = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,taxID)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,excludeCalculation)
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
				
	Public Sub modifyTaxsGroup(id,strDesc, objConn)
		on error resume next
		Dim strSQL

		strSQL = "UPDATE tax_group SET "
		strSQL = strSQL & "description=?"
		strSQL = strSQL & " WHERE id=?;"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
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
				
	Public Sub modifyTaxsGroupValue(idGroup,countryCode, stateRegionCode, taxID, excludeCalculation, objConn)
		on error resume next
		Dim strSQL

		strSQL = "UPDATE tax_group_value SET "
		if(isNull(taxID) OR taxID = "") then
			strSQL = strSQL & "id_tassa_applicata=NULL"
		else
			strSQL = strSQL & "id_tassa_applicata=?"			
		end if
		strSQL = strSQL & ", exclude_calculation=?"
		strSQL = strSQL & " WHERE id_group=? AND country_code=?"
		if not(isNull(stateRegionCode)) AND (stateRegionCode <> "") then
			strSQL = strSQL & " AND state_region_code=?"			
		else
			strSQL = strSQL & " AND ISNULL(state_region_code)"
		end if
		strSQL = strSQL & ";"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if not isNull(taxID) AND not(taxID = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,taxID)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,excludeCalculation)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode)
		if not isNull(stateRegionCode) AND not(stateRegionCode = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode)
		end if
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
		
	Public Sub deleteTaxsGroup(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM tax_group WHERE id=?;" 
		strSQL2 = "DELETE FROM tax_group_value WHERE id_group=?;"

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
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,3,1,,id)
		objConn.BeginTrans

		if(Application("use_innodb_table") = 0) then
			objCommand2.Execute()
		end if
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
		
	Public Sub deleteTaxsGroupValue(idGroup,countryCode, stateRegionCode)
		on error resume next
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "DELETE FROM tax_group_value"
		strSQL = strSQL & " WHERE id_group=? AND country_code=?"
		if not(isNull(stateRegionCode)) AND (stateRegionCode <> "") then
			strSQL = strSQL & " AND state_region_code=?"			
		else
			strSQL = strSQL & " AND ISNULL(state_region_code)"
		end if
		strSQL = strSQL & ";" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idGroup)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,countryCode)
		if not isNull(stateRegionCode) AND not(stateRegionCode = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,stateRegionCode)
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