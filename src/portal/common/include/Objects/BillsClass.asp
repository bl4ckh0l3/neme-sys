<%

Class BillsClass
	
	Private id
	Private descrizione_spesa
	Private valore
	Private tipologia_valore
	Private tassa_applicata
	Private applica_frontend
	Private applica_backend
	Private autoactive
	Private multiply
	Private required
	Private group
	Private taxsGroup
	Private type_view

	Private id_conf
	Private id_spesa_conf
	Private id_prod_field_conf
	Private rate_from_conf
	Private rate_to_conf
	Private operation_conf
	Private valore_conf
	
	
	Public Function getSpeseID()
		getSpeseID = id
	End Function
	
	Public Sub setSpeseID(strID)
		id = strID
	End Sub	
	
	Public Function getDescrizioneSpesa()
		getDescrizioneSpesa = descrizione_spesa
	End Function
	
	Public Sub setDescrizioneSpesa(strDesc)
		descrizione_spesa = strDesc
	End Sub		
	
	Public Function getValore()
		getValore = Cdbl(valore)
	End Function
	
	Public Sub setValore(strValore)
		valore = strValore
	End Sub
	
	
	Public Function getTipoValore()
		getTipoValore = tipologia_valore
	End Function
	
	Public Sub setTipoValore(strTipoValore)
		tipologia_valore = strTipoValore
	End Sub	

	Public Function getIDTassaApplicata()
		getIDTassaApplicata = tassa_applicata
	End Function
	
	Public Sub setIDTassaApplicata(strIDTassaApplicata)
		tassa_applicata = strIDTassaApplicata
	End Sub
	
	Public Function getApplicaFrontend()
		getApplicaFrontend = applica_frontend
	End Function
	
	Public Sub setApplicaFrontend(strApplicaFrontend)
		applica_frontend = strApplicaFrontend
	End Sub	
	
	Public Function getApplicaBackend()
		getApplicaBackend = applica_backend
	End Function
	
	Public Sub setApplicaBackend(strApplicaBackend)
		applica_backend = strApplicaBackend
	End Sub	
	
	Public Function getAutoactive()
		getAutoactive = autoactive
	End Function
	
	Public Sub setAutoactive(strAutoactive)
		autoactive = strAutoactive
	End Sub	
	
	Public Function getMultiply()
		getMultiply = multiply
	End Function
	
	Public Sub setMultiply(strMultiply)
		multiply = strMultiply
	End Sub	
	
	Public Function getRequired()
		getRequired = required
	End Function
	
	Public Sub setRequired(strRequired)
		required = strRequired
	End Sub		
	
	Public Function getGroup()
		getGroup = group
	End Function
	
	Public Sub setGroup(strGroup)
		group = strGroup
	End Sub		
	
	Public Function getTaxGroup()
		getTaxGroup = taxsGroup
	End Function
	
	Public Sub setTaxGroup(strTaxGroup)
		taxsGroup = strTaxGroup
	End Sub		
	
	Public Function getTypeView()
		getTypeView = type_view
	End Function
	
	Public Sub setTypeView(strTypeView)
		type_view = strTypeView
	End Sub
	
	
	Public Function getTaxGroupObj(iTaxGroup)
		getTaxGroupObj = null
		if (not(isNull(iTaxGroup)) AND iTaxGroup<>"")then
			On Error Resume Next
			Set objTG = New TaxsGroupClass
			Set getTaxGroupObj = objTG.getGroupByID(iTaxGroup)			
			Set objTG = nothing
			if(Err.number <> 0) then
				Set getTaxGroupObj = null
			end if
		end if
	End Function


	'************ metodi GET e SET per le spese accessorie config
	Public Function getConfID()
		getConfID = id_conf
	End Function
	
	Public Sub setConfID(strIDC)
		id_conf = strIDC
	End Sub	

	Public Function getSpeseConfID()
		getSpeseConfID = id_spesa_conf
	End Function
	
	Public Sub setSpeseConfID(strIDConf)
		id_spesa_conf = strIDConf
	End Sub	

	Public Function getProdFieldConfID()
		getProdFieldConfID = id_prod_field_conf
	End Function
	
	Public Sub setProdFieldConfID(strIDPFConf)
		id_prod_field_conf = strIDPFConf
	End Sub	

	Public Function getRateFromConf()
		getRateFromConf = rate_from_conf
	End Function
	
	Public Sub setRateFromConf(strRateC)
		rate_from_conf = strRateC
	End Sub

	Public Function getRateToConf()
		getRateToConf = rate_to_conf
	End Function
	
	Public Sub setRateToConf(strRatetC)
		rate_to_conf = strRatetC
	End Sub	

	Public Function getOperationConf()
		getOperationConf = operation_conf
	End Function
	
	Public Sub setOperationConf(strOperationC)
		operation_conf = strOperationC
	End Sub		

	Public Function getValoreConf()
		getValoreConf = valore_conf
	End Function
	
	Public Sub setValoreConf(strValoreC)
		valore_conf = strValoreC
	End Sub	


	Public Function getImpByStrategy(totaleProdottoImp4spese, totQta, obiListfieldProd)
		getImpByStrategy = 0
		On Error Resume Next

		'response.write("totaleProdottoImp4spese: "&totaleProdottoImp4spese&"<br>")
		'response.write("totQta: "&totQta&"<br>")
		'response.write("obiListfieldProd.count: "&obiListfieldProd.count&"<br>")
		'response.write("getTipoValore(): "&getTipoValore()&"<br>")
		'response.write("getValore(): "&getValore()&"<br>")

		Select Case CInt(getTipoValore())
			Case 1
				getImpByStrategy = CDbl(getValore())				
			Case 2
				getImpByStrategy = CDbl(totaleProdottoImp4spese) / 100 * CDbl(getValore())				
			Case 3
				On Error Resume Next
				Set objListBillsConf = getListaSpeseConfig(getSpeseID(), null)
				'response.write("objListBillsConf.count:"&objListBillsConf.count&"<br>")
				for each g in objListBillsConf
					tmpRF = objListBillsConf(g).getRateFromConf()
					tmpRT = objListBillsConf(g).getRateToConf()
					'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
					if(CDbl(totaleProdottoImp4spese)>=CDbl(tmpRF) AND CDbl(totaleProdottoImp4spese)<=CDbl(tmpRT))then
						getImpByStrategy = CDbl(objListBillsConf(g).getValoreConf())
						Exit for
					end if
				next
				Set objListBillsConf = nothing
				if(Err.number <> 0) then
					getImpByStrategy = 0
				end if				
			Case 4
				On Error Resume Next
				Set objListBillsConf = getListaSpeseConfig(getSpeseID(), null)
				'response.write("objListBillsConf.count:"&objListBillsConf.count&"<br>")
				for each g in objListBillsConf
					tmpRF = objListBillsConf(g).getRateFromConf()
					tmpRT = objListBillsConf(g).getRateToConf()
					'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
					if(CDbl(totaleProdottoImp4spese)>=CDbl(tmpRF) AND CDbl(totaleProdottoImp4spese)<=CDbl(tmpRT))then
						getImpByStrategy = CDbl(totaleProdottoImp4spese) / 100 * CDbl(objListBillsConf(g).getValoreConf())
						Exit for
					end if
				next
				Set objListBillsConf = nothing
				if(Err.number <> 0) then
					getImpByStrategy = 0
				end if				
			Case 5
				On Error Resume Next
				Set objListBillsConf = getListaSpeseConfig(getSpeseID(), null)
				'response.write("objListBillsConf.count:"&objListBillsConf.count&"<br>")
				for each g in objListBillsConf
					tmpRF = objListBillsConf(g).getRateFromConf()
					tmpRT = objListBillsConf(g).getRateToConf()
					'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
					if(CDbl(totQta)>=CDbl(tmpRF) AND CDbl(totQta)<=CDbl(tmpRT))then
						getImpByStrategy = CDbl(objListBillsConf(g).getValoreConf())
						Exit for
					end if
				next
				Set objListBillsConf = nothing
				if(Err.number <> 0) then
					getImpByStrategy = 0
				end if					
			Case 6
				On Error Resume Next
				Set objListBillsConf = getListaSpeseConfig(getSpeseID(), null)
				'response.write("objListBillsConf.count:"&objListBillsConf.count&"<br>")
				'response.write("totQta:"&totQta&"<br>")
				for each g in objListBillsConf
					tmpRF = objListBillsConf(g).getRateFromConf()
					if(CDbl(tmpRF)>CDbl(totQta))then Exit for
					tmpRT = objListBillsConf(g).getRateToConf()
					if(CDbl(tmpRT)>CDbl(totQta))then tmpRT=CDbl(totQta)
					'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
					for counterR=Cint(tmpRF) to Cint(tmpRT)
					'response.write("counter:"&counterR&"<br>")
						Select Case CInt(objListBillsConf(g).getOperationConf())
							Case 1
							getImpByStrategy = getImpByStrategy+CDbl(objListBillsConf(g).getValoreConf())
							Case 2
							getImpByStrategy = getImpByStrategy-CDbl(objListBillsConf(g).getValoreConf())
						End Select
					'response.write("getImpByStrategy t: "&getImpByStrategy&"<br>")
					next
				next
				Set objListBillsConf = nothing
				if(Err.number <> 0) then
					getImpByStrategy = 0
				end if			
			Case 7
				On Error Resume Next
				Set objListBillsConf = getListaSpeseConfig(getSpeseID(), null)
				'response.write("objListBillsConf.count:"&objListBillsConf.count&"<br>")

				for each g in objListBillsConf
					tmpIFP = objListBillsConf(g).getProdFieldConfID()
					tmpRF = objListBillsConf(g).getRateFromConf()
					tmpRT = objListBillsConf(g).getRateToConf()
					'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"; - tmpIFP:"&tmpIFP&"<br>")

					for each l in obiListfieldProd
						'response.write("l: "&x&"<br>")
						'response.write("obiListfieldProd(l).count: "&obiListfieldProd(l).count&"<br>")
						for each p in obiListfieldProd(l)
							if(p("id")=tmpIFP)then
								'response.write("p.count: "&p.count&"<br>")
								'response.write("p(id): "&p("id")&"<br>")
								'response.write("p(value): "&p("value")&"<br>")
								'response.write("p(qta): "&p("qta")&"<br>")
								if(CDbl(p("value"))>=CDbl(tmpRF) AND CDbl(p("value"))<=CDbl(tmpRT))then
									getImpByStrategy = getImpByStrategy+CDbl(objListBillsConf(g).getValoreConf())
									'response.write("getImpByStrategy: "&getImpByStrategy&"<br>")
								end if
							end if
						next
					next
				next
				Set objListBillsConf = nothing
				if(Err.number <> 0) then
					getImpByStrategy = 0
				end if				
			Case 8
				On Error Resume Next
				Set objListBillsConf = getListaSpeseConfig(getSpeseID(), null)
				'response.write("objListBillsConf.count:"&objListBillsConf.count&"<br>")

				for each g in objListBillsConf
					tmpIFP = objListBillsConf(g).getProdFieldConfID()
					tmpRF = objListBillsConf(g).getRateFromConf()
					tmpRT = objListBillsConf(g).getRateToConf()
					tmpTotVal = 0
					bolDoNext = true

					'qui devo fare due volte il doppio for, la prima per fare la somma dei value per l'id prodotto selezionato
					'la seconda per applicare il calcolo
					for each l in obiListfieldProd
						'response.write("l: "&x&"<br>")
						'response.write("obiListfieldProd(l).count: "&obiListfieldProd(l).count&"<br>")
						for each p in obiListfieldProd(l)
							'response.write("p.count: "&p.count&"<br>")
							'response.write("p(id): "&p("id")&"<br>")
							'response.write("p(value): "&p("value")&"<br>")
							'response.write("p(qta): "&p("qta")&"<br>")
							'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"; - tmpIFP:"&tmpIFP&"<br>")
							if(p("id")=tmpIFP)then
								tmpTotVal = tmpTotVal+(CDbl(p("value"))*CDbl(p("qta")))
								'response.write("tmpTotVal: "&tmpTotVal&"<br>")
							end if
						next
					next

					if(CDbl(tmpRF)>CDbl(tmpTotVal))then bolDoNext = false
					if(bolDoNext)then
						if(CDbl(tmpRT)>CDbl(tmpTotVal))then tmpRT=CDbl(tmpTotVal)

						'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
						for counterR=Cint(tmpRF) to Cint(tmpRT)
							'response.write("counter:"&counterR&"<br>")
							Select Case CInt(objListBillsConf(g).getOperationConf())
								Case 1
								getImpByStrategy = getImpByStrategy+CDbl(objListBillsConf(g).getValoreConf())
								Case 2
								getImpByStrategy = getImpByStrategy-CDbl(objListBillsConf(g).getValoreConf())
							End Select
							'response.write("getImpByStrategy t: "&getImpByStrategy&"<br>")
						next
					end if
				next
				Set objListBillsConf = nothing
				if(Err.number <> 0) then
					getImpByStrategy = 0
				end if	
		End Select	

		if(Err.number <> 0) then
			getImpByStrategy = 0
		end if	
		'response.write("getImpByStrategy: "&getImpByStrategy&"<br>")
	End Function	
	
	Public Function getListaSpese(descrizione_spesa, tipologia_valore, activeFront, activeBack)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getListaSpese = null		
		strSQL = "SELECT * FROM spese_accessorie"
		
		if (isNull(descrizione_spesa) AND isNull(tipologia_valore) AND isNull(activeFront) AND isNull(activeBack)) then
			strSQL = "SELECT * FROM spese_accessorie"
		else
			strSQL = strSQL & " WHERE"
			if not(isNull(descrizione_spesa)) then strSQL = strSQL & " AND descrizione_spesa=?"
			if not(isNull(tipologia_valore)) then strSQL = strSQL & " AND tipologia_valore=?"
			if not(isNull(activeFront)) then strSQL = strSQL & " AND applica_frontend=?"
			if not(isNull(activeBack)) then strSQL = strSQL & " AND applica_backend=?"
		end if
		
		strSQL = strSQL & " ORDER BY autoactive DESC, `group` ASC, descrizione_spesa DESC;"
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if (isNull(descrizione_spesa) AND isNull(tipologia_valore) AND isNull(activeFront) AND isNull(activeBack)) then
		else
			if not(isNull(descrizione_spesa)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,descrizione_spesa)
			if not(isNull(tipologia_valore)) then objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,tipologia_valore)
			if not(isNull(activeFront)) then objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activeFront)
			if not(isNull(activeBack)) then objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activeBack)			
		end if
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objSpese
			do while not objRS.EOF				
				Set objSpese = new BillsClass
				strID = objRS("id")
				objSpese.setSpeseID(strID)
				objSpese.setDescrizioneSpesa(objRS("descrizione_spesa"))
				objSpese.setValore(objRS("valore"))	
				objSpese.setTipoValore(objRS("tipologia_valore"))	
				objSpese.setIDTassaApplicata(objRS("id_tassa_applicata"))	
				objSpese.setApplicaFrontend(objRS("applica_frontend"))	
				objSpese.setApplicaBackend(objRS("applica_backend"))		
				objSpese.setAutoactive(objRS("autoactive"))		
				objSpese.setMultiply(objRS("multiply"))			
				objSpese.setRequired(objRS("required"))	
				objSpese.setGroup(objRS("group"))		
				objSpese.setTaxGroup(objRS("taxs_group"))
				objSpese.setTypeView(objRS("type_view"))
				objDict.add strID, objSpese
				objRS.moveNext()
			loop
			Set objSpese = nothing							
			Set getListaSpese = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	
	
	Public Function findSpesaByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		findSpesaByID = null		
		strSQL = "SELECT * FROM spese_accessorie WHERE id =?;"
		strSQL = Trim(strSQL)
		
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
			Dim objSpese		
			Set objSpese = new BillsClass
			strID = objRS("id")
			objSpese.setSpeseID(strID)
			objSpese.setDescrizioneSpesa(objRS("descrizione_spesa"))
			objSpese.setValore(objRS("valore"))	
			objSpese.setTipoValore(objRS("tipologia_valore"))	
			objSpese.setIDTassaApplicata(objRS("id_tassa_applicata"))
			objSpese.setApplicaFrontend(objRS("applica_frontend"))	
			objSpese.setApplicaBackend(objRS("applica_backend"))			
			objSpese.setAutoactive(objRS("autoactive"))		
			objSpese.setMultiply(objRS("multiply"))			
			objSpese.setRequired(objRS("required"))	
			objSpese.setGroup(objRS("group"))
			objSpese.setTaxGroup(objRS("taxs_group"))	
			objSpese.setTypeView(objRS("type_view"))				
			Set findSpesaByID = objSpese			
			Set objSpese = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function insertSpesa(descrizione_spesa, valore, tipologia_valore, id_tassa_applicata, applica_frontend, applica_backend, autoactive, multiply, required, group, tax_group, type_view, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		insertSpesa = -1
		
		strSQL = "INSERT INTO spese_accessorie(descrizione_spesa, valore, tipologia_valore, id_tassa_applicata, applica_frontend, applica_backend, autoactive, `multiply`, required, `group`, taxs_group, `type_view`) VALUES("
		strSQL = strSQL & "?,?,?,"
		if(isNull(id_tassa_applicata) OR id_tassa_applicata = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if		
		strSQL = strSQL & "?,?,?,?,?,?,"
		if(isNull(tax_group) OR tax_group = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		strSQL = strSQL & "?);"
							
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,descrizione_spesa)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,tipologia_valore)
		if not isNull(id_tassa_applicata) AND not(id_tassa_applicata = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_tassa_applicata)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,applica_frontend)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,applica_backend)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,autoactive)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,multiply)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,group)
		if not isNull(tax_group) AND not(tax_group = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,tax_group)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,type_view)
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(spese_accessorie.id) as id FROM spese_accessorie")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertSpesa = objRS("id")	
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
		
	Public Sub modifySpesa(id, descrizione_spesa, valore, tipologia_valore, id_tassa_applicata, applica_frontend, applica_backend, autoactive, multiply, required, group, tax_group, type_view, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE spese_accessorie SET "
		strSQL = strSQL & "id=?,"
		strSQL = strSQL & "descrizione_spesa=?,"
		strSQL = strSQL & "valore=?,"
		strSQL = strSQL & "tipologia_valore=?,"
		if(isNull(id_tassa_applicata) OR id_tassa_applicata = "") then
			strSQL = strSQL & "id_tassa_applicata=NULL,"
		else
			strSQL = strSQL & "id_tassa_applicata=?,"			
		end if
		strSQL = strSQL & "applica_frontend=?,"
		strSQL = strSQL & "applica_backend=?,"
		strSQL = strSQL & "autoactive=?,"
		strSQL = strSQL & "`multiply`=?,"
		strSQL = strSQL & "`required`=?,"
		strSQL = strSQL & "`group`=?,"
		if(isNull(tax_group) OR tax_group = "") then
			strSQL = strSQL & "taxs_group=NULL,"
		else
			strSQL = strSQL & "taxs_group=?,"			
		end if
		strSQL = strSQL & "`type_view`=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,descrizione_spesa)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,tipologia_valore)
		if not isNull(id_tassa_applicata) AND not(id_tassa_applicata = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_tassa_applicata)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,applica_frontend)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,applica_backend)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,autoactive)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,multiply)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,required)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,group)
		if not isNull(tax_group) AND not(tax_group = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,tax_group)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,type_view)
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
		
	Public Sub deleteSpesa(id)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL2 = "DELETE FROM spese_accessorie_config WHERE id_spesa=?;"
		strSQL = "DELETE FROM spese_accessorie WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
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
	End Sub

	Public Function findSpeseAssociations(id_spesa)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		Dim strSQL2, strSQL3, strSQL4
		findSpeseAssociations = false	
		strSQL = "SELECT * FROM spese_x_ordine WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_spesa)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then							
			findSpeseAssociations = true				
		end if	
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	
	Public Function getListaSpeseConfig(id_spesa, id_prod_field)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getListaSpeseConfig = null		
		strSQL = "SELECT * FROM spese_accessorie_config WHERE id_spesa=?"		
		if not(isNull(id_prod_field)) then strSQL = strSQL & " AND id_prod_field=?"		
		strSQL = strSQL & " ORDER BY id_prod_field, rate_from, rate_to ASC;"
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_spesa)
		if not(isNull(id_prod_field)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod_field)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objSpese
			do while not objRS.EOF				
				Set objSpese = new BillsClass
				strID = objRS("id")
				objSpese.setConfID(strID)
				objSpese.setSpeseConfID(objRS("id_spesa"))
				objSpese.setProdFieldConfID(objRS("id_prod_field"))	
				objSpese.setRateFromConf(objRS("rate_from"))
				objSpese.setRateToConf(objRS("rate_to"))
				objSpese.setOperationConf(objRS("operation"))	
				objSpese.setValoreConf(objRS("valore"))
				objDict.add strID, objSpese
				objRS.moveNext()
			loop
			Set objSpese = nothing							
			Set getListaSpeseConfig = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findSpesaConfigByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		findSpesaConfigByID = null		
		strSQL = "SELECT * FROM spese_accessorie_config WHERE id =?;"
		strSQL = Trim(strSQL)
		
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
			Dim objSpese		
			Set objSpese = new BillsClass
			strID = objRS("id")
			objSpese.setConfID(strID)
			objSpese.setSpeseConfID(objRS("id_spesa"))
			objSpese.setProdFieldConfID(objRS("id_prod_field"))	
			objSpese.setRateFromConf(objRS("rate_from"))
			objSpese.setRateToConf(objRS("rate_to"))
			objSpese.setOperationConf(objRS("operation"))	
			objSpese.setValoreConf(objRS("valore"))								
			Set findSpesaConfigByID = objSpese			
			Set objSpese = nothing			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

		
	Public Sub insertSpesaConfig(id_spesa, id_prod_field, rate_from, rate_to, operation, valore, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO spese_accessorie_config(id_spesa, id_prod_field, rate_from, rate_to, operation, `valore`) VALUES("
		strSQL = strSQL & "?,"
		if(isNull(id_prod_field) OR id_prod_field = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if		
		strSQL = strSQL & "?,?,"
		if(isNull(operation) OR operation = "") then
			strSQL = strSQL & "0"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ",?);"
						
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		if not isNull(id_prod_field) AND not(id_prod_field = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_field)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_from))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_to))
		if not isNull(operation) AND not(operation = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,operation)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
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
		
	Public Sub modifySpesaConfig(id, id_spesa, id_prod_field, rate_from, rate_to, operation, valore, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE spese_accessorie_config SET "
		strSQL = strSQL & "id_spesa=?,"
		if(isNull(id_prod_field) OR id_prod_field = "") then
			strSQL = strSQL & "id_prod_field=NULL,"
		else
			strSQL = strSQL & "id_prod_field=?,"			
		end if
		strSQL = strSQL & "rate_from=?,"
		strSQL = strSQL & "rate_to=?,"
		if(isNull(operation) OR operation = "") then
			strSQL = strSQL & "operation=0,"
		else
			strSQL = strSQL & "operation=?,"			
		end if
		strSQL = strSQL & "`valore`=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		if not isNull(id_prod_field) AND not(id_prod_field = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_field)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_from))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_to))
		if not isNull(operation) AND not(operation = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,operation)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
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
		
	Public Sub deleteSpesaConfigNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM spese_accessorie_config WHERE id=?;"
		
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
		
	Public Sub deleteSpesaConfig(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM spese_accessorie_config WHERE id=?;"
		
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
		
	Public Sub deleteSpesaConfigBySpesaNoTransaction(id_spesa)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM spese_accessorie_config WHERE id_spesa=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteSpesaConfigBySpesa(id_spesa, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM spese_accessorie_config WHERE id_spesa=?;"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_spesa)
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
	
	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = doubleValue
		
		'if (Application("dbType") = 0) then
			convertDoubleDelimiter = Replace(convertDoubleDelimiter, ".",",")
		'else		
			'convertDoubleDelimiter = Replace(convertDoubleDelimiter, ",",".")
		'end if			
	End Function
End Class
%>