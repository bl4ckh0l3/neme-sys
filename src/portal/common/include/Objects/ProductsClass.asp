<%

Class ProductsClass
	Private idProdotto
	Private nomeProdotto
	Private sommarioProd
	Private descProdotto
	Private prezzo
	Private qta_disp
	Private attivo
	Private sconto
	Private objFileProdotto
	Private objTargetProdotto
	Private codiceProd
	Private idTassaApplicata
	Private prod_type
	Private maxDownload
	Private maxDownloadTime
	Private descCatRelProd
	Private taxsGroup
	Private meta_description
	Private meta_keyword
	Private page_title
	Private edit_buy_qta
	
	
	Public Function getIDProdotto()
		getIDProdotto = idProdotto
	End Function
		
	Public Function getNomeProdotto()
		getNomeProdotto = nomeProdotto
	End Function
		
	Public Function getSommarioProdotto()
		getSommarioProdotto = sommarioProd
	End Function
	
	Public Function getDescProdotto()
		getDescProdotto = descProdotto
	End Function
	
	Public Function getPrezzo()
		getPrezzo = CDbl(prezzo)
	End Function
	
	Public Function getPrezzoScontato()
		getPrezzoScontato = CDbl(prezzo) - (CDbl(prezzo) / 100 * CDbl(sconto))
	End Function
	
	Public Function getImportoTassa(dblPrezzo)
		dim objTassa, importo, iValore, objTassaTmp
		Set objTassa = new TaxsClass
		Set objTassaTmp = objTassa.findTassaByID(idTassaApplicata)		
		
		iValore = objTassaTmp.getValore()
		iValore = CDbl(iValore)
		if(objTassaTmp.getTipoValore() = 2) then
			importo = CDbl(dblPrezzo) * (iValore / 100)
		else
			importo = iValore
		end if
		
		getImportoTassa = importo
		Set objTassaTmp = nothing
		Set objTassa = nothing
	End Function
	
	Public Function getQtaDisp()
		getQtaDisp = qta_disp
	End Function
	
	Public Function getAttivo()
		getAttivo = attivo
	End Function
	
	Public Function getSconto()
		getSconto = Cdbl(sconto)
	End Function
	
	Public Function hasSconto()
		hasSconto = Cdbl(sconto) > 0
	End Function

	Public Function getFileXProdotto()
		if(isNull(objFileProdotto) or isEmpty(objFileProdotto)) then
			getFileXProdotto = null
		else
			Set getFileXProdotto = objFileProdotto
		end if
	End Function
	
	Public Function getCodiceProd()
		getCodiceProd = codiceProd
	End Function
	
	Public Sub setFileXProdotto(objFiles)
		if(isNull(objFiles)) then
			objFileProdotto = null
		else
			Set objFileProdotto = objFiles
		end if		
	End Sub	
		
	Public Function getListaTarget()
		Set getListaTarget = objTargetProdotto
	End Function

	Public Function getIDTassaApplicata()
		getIDTassaApplicata = idTassaApplicata
	End Function
	
	Public Sub setListaTarget(objTarget)
		Set objTargetProdotto = objTarget
	End Sub
				
	Public Sub setIDProdotto(numIDProdotto)
		idProdotto = numIDProdotto
	End Sub
		
	Public Sub setNomeProdotto(strNomeProdotto)
		nomeProdotto = strNomeProdotto
	End Sub
		
	Public Sub setSommarioProdotto(strSommarioProdotto)
		sommarioProd = strSommarioProdotto
	End Sub
	
	Public Sub setDescProdotto(strDescProdotto)
		descProdotto = strDescProdotto
	End Sub
	
	Public Sub setPrezzo(strPrezzo)
		prezzo = strPrezzo
	End Sub
	
	Public Sub setQtaDisp(strQtaDisp)
		qta_disp = strQtaDisp
	End Sub
	
	Public Sub setAttivo(strAttivo)
		attivo = strAttivo
	End Sub
	
	Public Sub setSconto(strSconto)
		sconto = strSconto
	End Sub
	
	Public Sub setCodiceProd(strCodiceProd)
		codiceProd = strCodiceProd
	End Sub
	
	Public Sub setIDTassaApplicata(strIDTassaApplicata)
		idTassaApplicata = strIDTassaApplicata
	End Sub

	Public Function getProdType()
		getProdType = prod_type
	End Function
	
	Public Sub setProdType(intProdType)
		prod_type = intProdType
	End Sub
	
	Public Function getMaxDownload()
		getMaxDownload = maxDownload
	End Function
	
	Public Sub setMaxDownload(strMaxDownload)
		maxDownload = strMaxDownload
	End Sub
	
	Public Function getMaxDownloadTime()
		getMaxDownloadTime = maxDownloadTime
	End Function
	
	Public Sub setMaxDownloadTime(strMaxDownloadTime)
		maxDownloadTime = strMaxDownloadTime
	End Sub
	
	Public Function getDescCatRelProd()
		getDescCatRelProd = descCatRelProd
	End Function
	
	Public Sub setDescCatRelProd(strdescCatRelProd)
		descCatRelProd = strdescCatRelProd
	End Sub		
	
	Public Function getTaxGroup()
		getTaxGroup = taxsGroup
	End Function
	
	Public Sub setTaxGroup(strTaxGroup)
		taxsGroup = strTaxGroup
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
	
	Public Function getMetaDescription()
		getMetaDescription = meta_description
	End Function
	
	Public Sub setMetaDescription(strMetaDesc)
		meta_description = strMetaDesc
	End Sub
	
	Public Function getMetaKeyword()
		getMetaKeyword = meta_keyword
	End Function
	
	Public Sub setMetaKeyword(strMetaKeyword)
		meta_keyword = strMetaKeyword
	End Sub
	
	Public Function getPageTitle()
		getPageTitle = page_title
	End Function
	
	Public Sub setPageTitle(strPageTitle)
		page_title = strPageTitle
	End Sub
	
	Public Function getEditBuyQta()
		getEditBuyQta = edit_buy_qta
	End Function
	
	Public Sub setEditBuyQta(strEditBuyQta)
		edit_buy_qta = strEditBuyQta
	End Sub
	


'*********************************** METODI PRODOTTO *********************** 				
	Public Function insertProdotto(strNomeProd, strSommarioProd, strDescProd, numPrezzo, numQtaDisp, bolAttivo, numSconto, codiceProd, idTassaApplicata, prod_type, maxDownload, maxDownloadTime, tax_group, strMetaDesc, strMetaKey, strPageTitle, bolEditBuyQta, objConn)
		on error resume next
		insertProdotto = -1
		
		Dim objDB, strSQL, strSQLSelect, objRS
		
		strSQL = "INSERT INTO prodotti(nome_prod, sommario_prod, desc_prod, prezzo, qta_disp, attivo, sconto, codice_prod, id_tassa_applicata, prod_type, max_download, max_download_time, taxs_group, meta_description, meta_keyword, page_title, edit_buy_qta) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		
		if(isNull(idTassaApplicata) OR idTassaApplicata = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		
		strSQL = strSQL & "?,?,?,"
		if(isNull(tax_group) OR tax_group = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ",?,?,?,?);"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strNomeProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strSommarioProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strDescProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numPrezzo))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,numQtaDisp)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolAttivo)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numSconto))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,codiceProd)
		if not isNull(idTassaApplicata) AND not(idTassaApplicata = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idTassaApplicata)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,maxDownload)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,maxDownloadTime)
		if not isNull(tax_group) AND not(tax_group = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,tax_group)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolEditBuyQta)
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(prodotti.id_prodotto) as id FROM prodotti")
		if not (objRS.EOF) then
			'objRS.MoveFirst()
			insertProdotto = objRS("id")	
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
		
	Public Function insertProdottoNoTransaction(strNomeProd, strSommarioProd, strDescProd, numPrezzo, numQtaDisp, bolAttivo, numSconto, codiceProd, idTassaApplicata, prod_type, maxDownload, maxDownloadTime, tax_group, strMetaDesc, strMetaKey, strPageTitle, bolEditBuyQta)
		on error resume next
		insertProdottoNoTransaction = -1
		
		Dim objDB, strSQL, strSQLSelect, objRS, objConn
		
		strSQL = "INSERT INTO prodotti(nome_prod, sommario_prod, desc_prod, prezzo, qta_disp, attivo, sconto, codice_prod, id_tassa_applicata, prod_type, max_download, max_download_time, taxs_group, meta_description, meta_keyword, page_title, edit_buy_qta) VALUES('"
		strSQL = strSQL & "?,?,?,?,?,?,?,?,"
		
		if(isNull(idTassaApplicata) OR idTassaApplicata = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		
		strSQL = strSQL & "?,?,?,"
		if(isNull(tax_group) OR tax_group = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ",?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strNomeProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strSommarioProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strDescProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numPrezzo))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,numQtaDisp)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolAttivo)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numSconto))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,codiceProd)
		if not isNull(idTassaApplicata) AND not(idTassaApplicata = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idTassaApplicata)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,maxDownload)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,maxDownloadTime)
		if not isNull(tax_group) AND not(tax_group = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,tax_group)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolEditBuyQta)
		objCommand.Execute()
		
		Set objRS = objConn.Execute("SELECT max(prodotti.id_prodotto) as id FROM prodotti")
		if not (objRS.EOF) then
			insertProdottoNoTransaction = objRS("id")	
		end if	
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyProdotto(id, strNomeProd, strSommarioProd, strDescProd, numPrezzo, numQtaDisp, bolAttivo, numSconto, codiceProd, idTassaApplicata, prod_type, maxDownload, maxDownloadTime, tax_group, strMetaDesc, strMetaKey, strPageTitle, bolEditBuyQta, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
			
		strSQL = "UPDATE prodotti SET "
		strSQL = strSQL & "nome_prod=?,"
		strSQL = strSQL & "sommario_prod=?,"
		strSQL = strSQL & "desc_prod=?,"
		strSQL = strSQL & "prezzo=?,"
		strSQL = strSQL & "qta_disp=?," 
		strSQL = strSQL & "attivo=?,"		 
		strSQL = strSQL & "sconto=?,"		 
		strSQL = strSQL & "codice_prod=?,"
		if(isNull(idTassaApplicata) OR idTassaApplicata = "") then
			strSQL = strSQL & "id_tassa_applicata=NULL,"
		else
			strSQL = strSQL & "id_tassa_applicata=?,"			
		end if
		strSQL = strSQL & "prod_type=?,"
		strSQL = strSQL & "max_download=?,"
		strSQL = strSQL & "max_download_time=?,"
		if(isNull(tax_group) OR tax_group = "") then
			strSQL = strSQL & "taxs_group=NULL"
		else
			strSQL = strSQL & "taxs_group=?"			
		end if
		strSQL = strSQL & ",meta_description=?,"
		strSQL = strSQL & "meta_keyword=?,"
		strSQL = strSQL & "page_title=?,"
		strSQL = strSQL & "edit_buy_qta=?"
		strSQL = strSQL & " WHERE id_prodotto=?;"	

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strNomeProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strSommarioProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strDescProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numPrezzo))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,numQtaDisp)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolAttivo)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numSconto))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,codiceProd)
		if not isNull(idTassaApplicata) AND not(idTassaApplicata = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idTassaApplicata)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,maxDownload)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,maxDownloadTime)
		if not isNull(tax_group) AND not(tax_group = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,tax_group)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolEditBuyQta)
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
		
	Public Sub modifyProdottoNoTransaction(id, strNomeProd, strSommarioProd, strDescProd, numPrezzo, numQtaDisp, bolAttivo, numSconto, codiceProd, idTassaApplicata, prod_type, maxDownload, maxDownloadTime, tax_group, strMetaDesc, strMetaKey, strPageTitle, bolEditBuyQta)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "UPDATE prodotti SET "
		strSQL = strSQL & "nome_prod=?,"
		strSQL = strSQL & "sommario_prod=?,"
		strSQL = strSQL & "desc_prod=?,"
		strSQL = strSQL & "prezzo=?,"
		strSQL = strSQL & "qta_disp=?," 
		strSQL = strSQL & "attivo=?,"		 
		strSQL = strSQL & "sconto=?,"		 
		strSQL = strSQL & "codice_prod=?,"
		if(isNull(idTassaApplicata) OR idTassaApplicata = "") then
			strSQL = strSQL & "id_tassa_applicata=NULL,"
		else
			strSQL = strSQL & "id_tassa_applicata=?,"			
		end if
		strSQL = strSQL & "prod_type=?,"
		strSQL = strSQL & "max_download=?,"
		strSQL = strSQL & "max_download_time=?,"
		if(isNull(tax_group) OR tax_group = "") then
			strSQL = strSQL & "taxs_group=NULL"
		else
			strSQL = strSQL & "taxs_group=?"			
		end if
		strSQL = strSQL & ",meta_description=?,"
		strSQL = strSQL & "meta_keyword=?,"
		strSQL = strSQL & "page_title=?,"
		strSQL = strSQL & "edit_buy_qta=?"
		strSQL = strSQL & " WHERE id_prodotto=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strNomeProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strSommarioProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strDescProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numPrezzo))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,numQtaDisp)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,bolAttivo)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(numSconto))
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,codiceProd)
		if not isNull(idTassaApplicata) AND not(idTassaApplicata = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,idTassaApplicata)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,prod_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,maxDownload)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,maxDownloadTime)
		if not isNull(tax_group) AND not(tax_group = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,tax_group)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,bolEditBuyQta)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteProdotto(id, objConn)
		on error resume next
		Dim objDB, strSQLDelCommenti, strSQLDelFile, strSQLDelTarget, strSQLDelRelProd, strSQLDelProdotto, objRS, strSQLDelDownProd, strSQLDelFieldTransProd, strSQLDelLocaliz
		strSQLDelDownProd = "DELETE FROM downloadable_products WHERE id_product=?;"
		strSQLDelCommenti = "DELETE FROM commenti WHERE id_element=?;"
		strSQLDelFile = "DELETE FROM attach_x_prodotti WHERE id_prodotto=?;"
		strSQLDelTarget = "DELETE FROM target_x_prodotto WHERE id_prodotto=?;"
		strSQLDelRelProd = "DELETE FROM relation_x_prodotto WHERE id_prod=? OR id_prod_rel=?;"
		strSQLDelFieldTransProd = "DELETE FROM prodotto_main_field_translation WHERE id_prod=?;"
		strSQLDelProdotto = "DELETE FROM prodotti WHERE id_prodotto=?;"
		strSQLDelLocaliz = "DELETE FROM googlemap_localization WHERE id_element=? AND `type`=2;"

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
		objCommand.CommandText = strSQLDelDownProd
		objCommand2.CommandText = strSQLDelCommenti
		objCommand3.CommandText = strSQLDelFile
		objCommand4.CommandText = strSQLDelTarget
		objCommand5.CommandText = strSQLDelRelProd
		objCommand6.CommandText = strSQLDelFieldTransProd
		objCommand7.CommandText = strSQLDelProdotto
		objCommand8.CommandText = strSQLDelLocaliz
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)
		objCommand4.Parameters.Append objCommand4.CreateParameter(,19,1,,id)
		objCommand5.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand5.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand6.Parameters.Append objCommand6.CreateParameter(,19,1,,id)
		objCommand7.Parameters.Append objCommand7.CreateParameter(,19,1,,id)
		objCommand8.Parameters.Append objCommand8.CreateParameter(,19,1,,id)

		if(Application("use_innodb_table") = 0) then
			objCommand.Execute()
			objCommand2.Execute()
			objCommand3.Execute()
			objCommand4.Execute()
			objCommand5.Execute()
			objCommand6.Execute()
		end if
		objCommand8.Execute()
		objCommand7.Execute()

		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		Set objCommand4 = Nothing
		Set objCommand5 = Nothing
		Set objCommand6 = Nothing
		Set objCommand7 = Nothing
		Set objCommand8 = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub				
		
	Public Sub deleteProdottoNoTransaction(id)
		on error resume next
		Dim objDB, strSQLDelCommenti, strSQLDelFile, strSQLDelTarget, strSQLDelRelProd, strSQLDelProdotto, objRS, objConn, strSQLDelDownProd, strSQLDelFieldTransProd, strSQLDelLocaliz
		strSQLDelDownProd = "DELETE FROM downloadable_products WHERE id_product=?;"
		strSQLDelCommenti = "DELETE FROM commenti WHERE id_prodotto=?;"
		strSQLDelFile = "DELETE FROM attach_x_prodotti WHERE id_prodotto=?;"
		strSQLDelTarget = "DELETE FROM target_x_prodotto WHERE id_prodotto=?;"
		strSQLDelRelProd = "DELETE FROM relation_x_prodotto WHERE id_prod=? OR id_prod_rel=?;"
		strSQLDelFieldTransProd = "DELETE FROM prodotto_main_field_translation WHERE id_prod=?;"
		strSQLDelProdotto = "DELETE FROM prodotti WHERE id_prodotto=?;"
		strSQLDelLocaliz = "DELETE FROM googlemap_localization WHERE id_element=? AND `type`=2;"

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
		objCommand.CommandText = strSQLDelDownProd
		objCommand2.CommandText = strSQLDelCommenti
		objCommand3.CommandText = strSQLDelFile
		objCommand4.CommandText = strSQLDelTarget
		objCommand5.CommandText = strSQLDelRelProd
		objCommand6.CommandText = strSQLDelFieldTransProd
		objCommand7.CommandText = strSQLDelProdotto
		objCommand8.CommandText = strSQLDelLocaliz
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand3.Parameters.Append objCommand3.CreateParameter(,19,1,,id)
		objCommand4.Parameters.Append objCommand4.CreateParameter(,19,1,,id)
		objCommand5.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand5.Parameters.Append objCommand5.CreateParameter(,19,1,,id)
		objCommand6.Parameters.Append objCommand6.CreateParameter(,19,1,,id)
		objCommand7.Parameters.Append objCommand7.CreateParameter(,19,1,,id)
		objCommand8.Parameters.Append objCommand8.CreateParameter(,19,1,,id)
		
		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand.Execute()
			objCommand2.Execute()
			objCommand3.Execute()
			objCommand4.Execute()
			objCommand5.Execute()
			objCommand6.Execute()
		end if
		objCommand8.Execute()
		objCommand7.Execute()

		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objCommand3 = Nothing
		Set objCommand4 = Nothing
		Set objCommand5 = Nothing
		Set objCommand6 = Nothing
		Set objCommand7 = Nothing
		Set objCommand8 = Nothing
		
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

	Public Function changeQtaProdottoNoTransaction(id_prodotto, qta, oldQta)
		on error resume next
		changeQtaProdottoNoTransaction = -1
		
		Dim objDB, strSQL, objRS, newQta
		Dim objConn
		
		newQta = CLng(oldQta) - CLng(qta)			
		
		strSQL = "UPDATE prodotti SET "
		strSQL = strSQL & "qta_disp=?"
		if(newQta <= 0) then
			strSQL = strSQL & ",attivo=0"
		else
			strSQL = strSQL & ",attivo=1"
		end if		
		strSQL = strSQL & " WHERE id_prodotto=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,newQta)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Execute()	
		
		if(newQta = 0)then
			changeQtaProdottoNoTransaction = 0
		elseif(newQta > 0)then
			changeQtaProdottoNoTransaction = 1
		end if
						
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Function changeQtaProdotto(id_prodotto, qta, oldQta, objConn)
		on error resume next
		changeQtaProdotto = -1
		
		Dim objDB, strSQL, objRS, newQta		
		newQta = CLng(oldQta) - CLng(qta)			
		
		strSQL = "UPDATE prodotti SET "
		strSQL = strSQL & "qta_disp=?"
		if(newQta <= 0) then
			strSQL = strSQL & ",attivo=0"
		else
			strSQL = strSQL & ",attivo=1"
		end if		
		strSQL = strSQL & " WHERE id_prodotto=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,newQta)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Execute()
		Set objCommand = Nothing

		if(newQta = 0)then
			changeQtaProdotto = 0
		elseif(newQta > 0)then
			changeQtaProdotto = 1
		end if
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub changeStatoProdottoNoTransaction(id_prodotto, stato_prod)
		on error resume next
		
		Dim objDB, strSQL, objRS
		Dim objConn		
		
		strSQL = "UPDATE prodotti SET "
		strSQL = strSQL & "attivo=?"	
		strSQL = strSQL & " WHERE id_prodotto=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,stato_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prodotto)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Sub	
		
	Public Function getListaProdotti(order_by, attachments)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objProdotto, objListaFiles
		Dim objFiles, objListaTarget
		getListaProdotti = null  
		strSQL = "SELECT * FROM prodotti"
		
		if not(isNull(order_by)) then
			Select Case order_by
			Case 101
				strSQL = strSQL & " ORDER BY id_prodotto ASC"
			Case 102
				strSQL = strSQL & " ORDER BY id_prodotto DESC"
			Case 103
				strSQL = strSQL & " ORDER BY nome_prod ASC"
			Case 104
				strSQL = strSQL & " ORDER BY nome_prod DESC"
			Case 105
				strSQL = strSQL & " ORDER BY prezzo ASC"
			Case 106
				strSQL = strSQL & " ORDER BY prezzo DESC"
			Case 107
				strSQL = strSQL & " ORDER BY qta_disp ASC"
			Case 108
				strSQL = strSQL & " ORDER BY qta_disp DESC"
			Case 109
				strSQL = strSQL & " ORDER BY attivo ASC"	
			Case 110
				strSQL = strSQL & " ORDER BY attivo DESC"	
			Case 111
				strSQL = strSQL & " ORDER BY codice_prod ASC"	
			Case 112
				strSQL = strSQL & " ORDER BY codice_prod DESC"		
			Case Else
			End Select
		end if
		strSQL = strSQL & ";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Set objRS = objConn.Execute(strSQL)
		
		if not(objRS.EOF) then		   
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
			Set objFiles = new File4ProductsClass  
			do while not objRS.EOF
				Set objProdotto = new ProductsClass
				strID = objRS("id_prodotto")
								
				objProdotto.setIDProdotto(objRS("id_prodotto"))    
				objProdotto.setNomeProdotto(objRS("nome_prod"))
				objProdotto.setSommarioProdotto(objRS("sommario_prod"))
				objProdotto.setDescProdotto(objRS("desc_prod"))
				objProdotto.setPrezzo(objRS("prezzo"))
				objProdotto.setQtaDisp(objRS("qta_disp"))
				objProdotto.setAttivo(objRS("attivo"))
				objProdotto.setSconto(objRS("sconto"))
				objProdotto.setCodiceProd(objRS("codice_prod"))
				objProdotto.setIDTassaApplicata(objRS("id_tassa_applicata"))
				objProdotto.setProdType(objRS("prod_type"))
				objProdotto.setMaxDownload(objRS("max_download"))
				objProdotto.setMaxDownloadTime(objRS("max_download_time"))
				objProdotto.setTaxGroup(objRS("taxs_group"))		
				objProdotto.setMetaDescription(objRS("meta_description"))	
				objProdotto.setMetaKeyword(objRS("meta_keyword"))
				objProdotto.setPageTitle(objRS("page_title"))
				objProdotto.setEditBuyQta(objRS("edit_buy_qta"))
				
				Set objListaTarget = objProdotto.getTargetPerProdotto(strID)
				if not(isEmpty(objListaTarget)) then
					objProdotto.setListaTarget(objListaTarget)
					Set objListaTarget = nothing
				else
					Set objListaTarget = nothing
					response.Redirect(Application("baseroot")&Application("error_page")&"?error=020")
				end if	
							
				if(CBool(attachments)) then 
					Set objListaFiles = objFiles.getFilePerProdotto(strID)				
					if not(isEmpty(objListaFiles)) then
						objProdotto.setFileXProdotto(objListaFiles)
						Set objListaFiles = nothing
					else
						Set objListaFiles = nothing
					end if
				end if
									
				objDict.add strID, objProdotto
				Set objProdotto = nothing
				objRS.moveNext()
			loop
			
			Set objFiles = Nothing
			Set getListaProdotti = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function

	Public Function findProdotti(id_prodotto, codice_prod, nome_prod, sommario_prod, desc_prod, prezzo, qta_ins, prod_type, attivo, order_by, arrTargetProd, arrTargetLang, bolAddTarget, bolAddAttachment)
		on error resume next
				
		findProdotti = null
		
		Dim objDB, strSQL, objRS, objConn, objDict, objProdotto, objListaFiles
		Dim objFiles, strSQLTarget, objRSTargetProd, objRSTargetLang, objListTarget
		Dim hasTarget
		hasTarget = true

		Dim noTargetProd,noTargetlang
		noTargetProd = (isNull(arrTargetProd) OR not(strComp(typename(arrTargetProd), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0)) 

		Set objListTarget = Server.CreateObject("Scripting.Dictionary")
		if (noTargetProd) AND (noTargetlang) then
			hasTarget = false
		elseif (noTargetProd) OR (noTargetlang) then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=024")		
		end if
		

		strSQL = "SELECT * FROM prodotti"		
		if (isNull(id_prodotto) AND isNull(codice_prod) AND isNull(nome_prod) AND isNull(sommario_prod) AND isNull(desc_prod) AND isNull(prezzo) AND isNull(qta_ins) AND isNull(prod_type) AND (isNull(attivo) OR (attivo = 0)) AND not(hasTarget)) then
			strSQL = "SELECT * FROM prodotti"
		else
			strSQL = strSQL & " WHERE"

			if not(isNull(id_prodotto)) then strSQL = strSQL & " AND id_prodotto=?"
			if not(isNull(codice_prod)) then strSQL = strSQL & " AND codice_prod=?"
			if not(isNull(nome_prod)) then strSQL = strSQL & " AND nome_prod LIKE ?"
			if not(isNull(sommario_prod)) then strSQL = strSQL & " AND sommario_prod LIKE ?"
			if not(isNull(desc_prod)) then strSQL = strSQL & " AND desc_prod LIKE ?"
			if not(isNull(prezzo)) then strSQL = strSQL & " AND prezzo =?"
			if not(isNull(qta_ins)) then strSQL = strSQL & " AND qta_ins =?"
			if not(isNull(prod_type)) then strSQL = strSQL & " AND prod_type IN("&prod_type&")"
			if (not(isNull(attivo)) AND (attivo = 1)) then strSQL = strSQL & " AND attivo=?"
			if (hasTarget) then 
				strSQL = strSQL & " AND id_prodotto IN("					
				strSQL = strSQL & "SELECT DISTINCT(id_prodotto) FROM target_x_prodotto WHERE id_prodotto IN("
				strSQL = strSQL & "SELECT DISTINCT(id_prodotto) FROM target_x_prodotto WHERE id_target IN("								
				for each idx in arrTargetProd
					strSQL = strSQL &idx&","
				next					
				strSQL = strSQL & "))"				
				strSQL = strSQL & " AND id_target IN("	
				for each idy in arrTargetLang
					strSQL = strSQL &idy&","
				next					
				strSQL = strSQL & "))"						
				strSQL = Replace(strSQL, ",)", ")", 1, -1, 1)
				strSQL = Trim(strSQL)
			end if
			
		end if
				
		if not(isNull(order_by)) then
			Select Case order_by
			Case 101
				strSQL = strSQL & " ORDER BY id_prodotto ASC"
			Case 102
				strSQL = strSQL & " ORDER BY id_prodotto DESC"
			Case 103
				strSQL = strSQL & " ORDER BY nome_prod ASC"
			Case 104
				strSQL = strSQL & " ORDER BY nome_prod DESC"
			Case 105
				strSQL = strSQL & " ORDER BY prezzo ASC"
			Case 106
				strSQL = strSQL & " ORDER BY prezzo DESC"
			Case 107
				strSQL = strSQL & " ORDER BY qta_disp ASC"
			Case 108
				strSQL = strSQL & " ORDER BY qta_disp DESC"
			Case 109
				strSQL = strSQL & " ORDER BY attivo ASC"	
			Case 110
				strSQL = strSQL & " ORDER BY attivo DESC"
			Case 111
				strSQL = strSQL & " ORDER BY codice_prod ASC"	
			Case 112
				strSQL = strSQL & " ORDER BY codice_prod DESC"			
			Case Else
			End Select
		end if
		
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"

	
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection() 
	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		if (isNull(id_prodotto) AND isNull(codice_prod) AND isNull(nome_prod) AND isNull(sommario_prod) AND isNull(desc_prod) AND isNull(prezzo) AND isNull(qta_ins) AND (isNull(attivo) OR (attivo = 0))) then
		else
			if not(isNull(id_prodotto)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prodotto)
			if not(isNull(codice_prod)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,codice_prod)
			if not(isNull(nome_prod)) then objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&nome_prod&"%")
			if not(isNull(sommario_prod)) then bjCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&sommario_prod&"%")
			if not(isNull(desc_prod)) then bjCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&desc_prod&"%")
			if not(isNull(prezzo)) then objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(prezzo))
			if not(isNull(qta_ins)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,qta_ins)
			if (not(isNull(attivo)) AND (attivo = 1)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,attivo)
		end if
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			findProdotti = null
		else  
			Dim objListaTarget, checkTarget
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			 
			do while not objRS.EOF
				Set objProdotto = new ProductsClass
				strID = objRS("id_prodotto")								
				objProdotto.setIDProdotto(objRS("id_prodotto"))    
				objProdotto.setNomeProdotto(objRS("nome_prod"))
				objProdotto.setSommarioProdotto(objRS("sommario_prod"))
				objProdotto.setDescProdotto(objRS("desc_prod"))
				objProdotto.setPrezzo(objRS("prezzo"))
				objProdotto.setQtaDisp(objRS("qta_disp"))
				objProdotto.setAttivo(objRS("attivo"))
				objProdotto.setSconto(objRS("sconto"))
				objProdotto.setCodiceProd(objRS("codice_prod"))
				objProdotto.setIDTassaApplicata(objRS("id_tassa_applicata"))
				objProdotto.setProdType(objRS("prod_type"))
				objProdotto.setMaxDownload(objRS("max_download"))
				objProdotto.setMaxDownloadTime(objRS("max_download_time"))
				objProdotto.setTaxGroup(objRS("taxs_group"))		
				objProdotto.setMetaDescription(objRS("meta_description"))	
				objProdotto.setMetaKeyword(objRS("meta_keyword"))
				objProdotto.setPageTitle(objRS("page_title"))
				objProdotto.setEditBuyQta(objRS("edit_buy_qta"))
													
				objDict.add strID, objProdotto
				Set objProdotto = nothing
				objRS.moveNext()
			loop

			if(CBool(bolAddTarget) OR CBool(bolAddAttachment))then
				Set objProdotto = new ProductsClass
				Set objFiles = new File4ProductsClass 

				for each j in objDict
					bolValid = true
					if(CBool(bolAddTarget))then				
						Set objListaTarget = objProdotto.getTargetPerProdotto(j)					
						if not(isEmpty(objListaTarget)) then
							objDict(j).setListaTarget(objListaTarget)
						else
							call objDict.remove(j)
							bolValid = false
						end if
						Set objListaTarget = nothing
					end if					
										
					if(bolValid AND CBool(bolAddAttachment)) then						
						on Error Resume Next				
						Set objListaFiles = objFiles.getFilePerProdotto(j)
						if Err.number <> 0 then
							objListaFiles = null
						end if
					
						if not(isNull(objListaFiles)) then
							objProdotto.setFileXProdotto(objListaFiles)
						end if
						Set objListaFiles = nothing
					end if				
				next

				Set objFiles = Nothing
				Set objProdotto = nothing
			end if
			
			if (objDict.Count > 0) then
				Set findProdotti = objDict
			else
				findProdotti = null			
			end if
						
			Set objDict = nothing    
		end if
		
		Set objListTarget = nothing
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function


	Public Function findProdottiCached(id_prodotto, codice_prod, nome_prod, sommario_prod, desc_prod, prezzo, qta_ins, prod_type, attivo, order_by, arrTargetProd, arrTargetLang, bolAddTarget, bolAddAttachment)				
		findProdottiCached = null 
		
		Dim objDB, strSQL, objRS, objConn, objDict, objProdotto, objListaFiles
		Dim objFiles, strSQLTarget, objRSTargetProd, objRSTargetLang, objListTarget
		Dim hasTarget, doExit
		hasTarget = true
		cacheKey="findp"

		Dim noTargetProd,noTargetlang
		noTargetProd = (isNull(arrTargetProd) OR not(strComp(typename(arrTargetProd), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0))

		if (noTargetProd) AND (noTargetlang) then
			hasTarget = false
		elseif (noTargetProd) OR (noTargetlang) then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=024")		
		end if
		
		strSQL = "SELECT * FROM prodotti"		
		if (isNull(id_prodotto) AND isNull(codice_prod) AND isNull(nome_prod) AND isNull(sommario_prod) AND isNull(desc_prod) AND isNull(prezzo) AND isNull(qta_ins) AND isNull(prod_type) AND (isNull(attivo) OR (attivo = 0)) AND not(hasTarget)) then
			strSQL = "SELECT * FROM prodotti"
		else
			strSQL = strSQL & " WHERE"
			
			Set objBase64 = new Base64Class

			if not(isNull(id_prodotto)) then 
				strSQL = strSQL & " AND id_prodotto=?"
				cacheKey=cacheKey&"-"&id_prodotto
			end if
			if not(isNull(codice_prod)) then 
				strSQL = strSQL & " AND codice_prod=?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(codice_prod)
			end if
			if not(isNull(nome_prod)) then 
				strSQL = strSQL & " AND nome_prod LIKE ?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(nome_prod)
			end if
			if not(isNull(sommario_prod)) then 
				strSQL = strSQL & " AND sommario_prod LIKE ?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(sommario_prod)
			end if
			if not(isNull(desc_prod)) then
				strSQL = strSQL & " AND desc_prod LIKE ?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(desc_prod)
			end if
			if not(isNull(prezzo)) then 
				strSQL = strSQL & " AND prezzo =?"
				cacheKey=cacheKey&"-"&objBase64.Base64Encode(prezzo)
			end if
			if not(isNull(qta_ins)) then 
				strSQL = strSQL & " AND qta_ins =?"
				cacheKey=cacheKey&"-"&qta_ins
			end if
			if not(isNull(prod_type)) then
				strSQL = strSQL & " AND prod_type IN("&prod_type&")"
				cacheKey=cacheKey&"-"&prod_type
			end if
			if (not(isNull(attivo)) AND (attivo = 1)) then
				strSQL = strSQL & " AND attivo=?"
				cacheKey=cacheKey&"-"&attivo
			end if
			if (hasTarget) then 
				strSQL = strSQL & " AND id_prodotto IN("					
				strSQL = strSQL & "SELECT DISTINCT(id_prodotto) FROM target_x_prodotto WHERE id_prodotto IN("
				strSQL = strSQL & "SELECT DISTINCT(id_prodotto) FROM target_x_prodotto WHERE id_target IN("								
				for each idx in arrTargetProd
					strSQL = strSQL &idx&","
					cacheKey=cacheKey&"-"&idx
				next					
				strSQL = strSQL & "))"				
				strSQL = strSQL & " AND id_target IN("	
				for each idy in arrTargetLang
					strSQL = strSQL &idy&","
					cacheKey=cacheKey&"-"&idy
				next					
				strSQL = strSQL & "))"						
				strSQL = Replace(strSQL, ",)", ")", 1, -1, 1)
				strSQL = Trim(strSQL)					
			end if

			Set objBase64 = nothing
		end if
				
		if not(isNull(order_by)) then
			Select Case order_by
			Case 101
				strSQL = strSQL & " ORDER BY id_prodotto ASC"
			Case 102
				strSQL = strSQL & " ORDER BY id_prodotto DESC"
			Case 103
				strSQL = strSQL & " ORDER BY nome_prod ASC"
			Case 104
				strSQL = strSQL & " ORDER BY nome_prod DESC"
			Case 105
				strSQL = strSQL & " ORDER BY prezzo ASC"
			Case 106
				strSQL = strSQL & " ORDER BY prezzo DESC"
			Case 107
				strSQL = strSQL & " ORDER BY qta_disp ASC"
			Case 108
				strSQL = strSQL & " ORDER BY qta_disp DESC"
			Case 109
				strSQL = strSQL & " ORDER BY attivo ASC"	
			Case 110
				strSQL = strSQL & " ORDER BY attivo DESC"
			Case 111
				strSQL = strSQL & " ORDER BY codice_prod ASC"	
			Case 112
				strSQL = strSQL & " ORDER BY codice_prod DESC"			
			Case Else
			End Select
			cacheKey=cacheKey&"-"&order_by
		end if
		
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"

		cacheKey=Trim(cacheKey)
		'response.write("cacheKey final: "&cacheKey&"<br>")
		
		'tento il recupero dell'oggetto dalla cache
		on error resume next
		Set ojbCache = new CacheClass
		
		Set cachedObj = ojbCache.getItem(cacheKey)
		
		'response.write("typename(cachedObj): "& typename(cachedObj)&"<br>")
		
		if (Instr(1, typename(cachedObj), "Dictionary", 1) > 0) then
			Set objListaProdC = Server.CreateObject("Scripting.Dictionary")

			'response.write("cachedObj.count: "& cachedObj.count&"<br>")
		
			for each skey in cachedObj
				Set objProdottoC = new ProductsClass						
				objProdottoC.setIDProdotto(cachedObj(skey)("id_prodotto"))    
				objProdottoC.setNomeProdotto(cachedObj(skey)("nome_prod"))
				objProdottoC.setSommarioProdotto(cachedObj(skey)("sommario_prod"))
				objProdottoC.setDescProdotto(cachedObj(skey)("desc_prod"))
				objProdottoC.setPrezzo(cachedObj(skey)("prezzo"))
				objProdottoC.setQtaDisp(cachedObj(skey)("qta_disp"))
				objProdottoC.setAttivo(cachedObj(skey)("attivo"))
				objProdottoC.setSconto(cachedObj(skey)("sconto"))
				objProdottoC.setCodiceProd(cachedObj(skey)("codice_prod"))
				objProdottoC.setIDTassaApplicata(cachedObj(skey)("id_tassa_applicata"))
				objProdottoC.setProdType(cachedObj(skey)("prod_type"))
				objProdottoC.setMaxDownload(cachedObj(skey)("max_download"))
				objProdottoC.setMaxDownloadTime(cachedObj(skey)("max_download_time"))
				objProdottoC.setTaxGroup(cachedObj(skey)("taxs_group"))		
				objProdottoC.setMetaDescription(cachedObj(skey)("meta_description"))	
				objProdottoC.setMetaKeyword(cachedObj(skey)("meta_keyword"))
				objProdottoC.setPageTitle(cachedObj(skey)("page_title"))
				objProdottoC.setEditBuyQta(cachedObj(skey)("edit_buy_qta"))
				
				'response.write("titolo:"&objProdottoC.getNomeProdotto()&" - id:"&objProdottoC.getIDProdotto()&"<br>")
				
				Set objListaTarget = Server.CreateObject("Scripting.Dictionary")				
				if (Instr(1, typename(cachedObj(skey)("target_list")), "Dictionary", 1) > 0) then
					Set objListaTargetTmp = cachedObj(skey)("target_list")
					for each xt in objListaTargetTmp
						Set objTarget = new Targetclass
						objTarget.setTargetID(xt)
						objTarget.setTargetDescrizione(objListaTargetTmp(xt)("descrizione"))
						objTarget.setTargetType(objListaTargetTmp(xt)("type"))	
						objListaTarget.add xt, objTarget
						'response.write("targetid:"&objTarget.getTargetID()&" - descrizione:"&objTarget.getTargetDescrizione()&"<br>")
						Set objTarget = nothing		
					next				
				
					Set objListaTargetTmp = nothing			
					objProdottoC.setListaTarget(objListaTarget)
					'response.write("objProdottoC.getListaTarget().count: "& objProdottoC.getListaTarget().count&"<br>")
					Set objListaTarget = nothing
				end if			
				
				Set objListaFiles = Server.CreateObject("Scripting.Dictionary")				
				if (Instr(1, typename(cachedObj(skey)("file_list")), "Dictionary", 1) > 0) then
					Set objListaFilesTmp = cachedObj(skey)("file_list")					
					for each xf in objListaFilesTmp
						Set objFiles = new File4ProductsClass
						objFiles.setFileID(xf)
						objFiles.setProdottoID(objListaFilesTmp(xf)("id_prodotto"))
						objFiles.setFileName(objListaFilesTmp(xf)("filename"))
						objFiles.setFileType(objListaFilesTmp(xf)("content_type"))
						objFiles.setFilePath(objListaFilesTmp(xf)("path"))
						objFiles.setFileDida(objListaFilesTmp(xf)("file_dida"))
						objFiles.setFileTypeLabel(objListaFilesTmp(xf)("file_label"))								
						objListaFiles.add xf, objFiles
						Set objFiles = nothing			
					next
					Set objListaFilesTmp = nothing
				end if
				
				if(objListaFiles.count>0)then
					objProdottoC.setFileXProdotto(objListaFiles)
				else
					objProdottoC.setFileXProdotto(null)
				end if
				Set objListaFiles = nothing
								
				'response.write("typename(objProdottoC): "& typename(objProdottoC)&"<br>")
			
				objListaProdC.add skey, objProdottoC
				Set objProdottoC = nothing	
			next	
			'response.write("objListaProdC.count: "& objListaProdC.count&"<br>")
			
			Set findProdottiCached = objListaProdC
			Set objListaProdC = nothing			
		else
			findProdottiCached = null
		end if
		
		if Err.number <> 0 then
			findProdottiCached = null
			'response.write(Err.number&" - "&Err.description&"<br>")
		end if
		
		'response.write("typename(findProdottiCached): "& typename(findProdottiCached)&"<br>")

		if not(Instr(1, typename(findProdottiCached), "Dictionary", 1) > 0) then
			on error resume next

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()  
		
			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQL

			if (isNull(id_prodotto) AND isNull(codice_prod) AND isNull(nome_prod) AND isNull(sommario_prod) AND isNull(desc_prod) AND isNull(prezzo) AND isNull(qta_ins) AND (isNull(attivo) OR (attivo = 0))) then
			else
				if not(isNull(id_prodotto)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prodotto)
				if not(isNull(codice_prod)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,codice_prod)
				if not(isNull(nome_prod)) then objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&nome_prod&"%")
				if not(isNull(sommario_prod)) then bjCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&sommario_prod&"%")
				if not(isNull(desc_prod)) then bjCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&desc_prod&"%")
				if not(isNull(prezzo)) then objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(prezzo))
				if not(isNull(qta_ins)) then objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,qta_ins)
				if (not(isNull(attivo)) AND (attivo = 1)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,attivo)
			end if
			Set objRS = objCommand.Execute()

			if objRS.EOF then
				findProdottiCached = null
			else  
				Dim objListaTarget, checkTarget
				Set objDict = Server.CreateObject("Scripting.Dictionary")
				Set objListaProdCache = Server.CreateObject("Scripting.Dictionary")
				 
				do while not objRS.EOF
					Set objProdotto = new ProductsClass
					Set objProdottoCache = Server.CreateObject("Scripting.Dictionary")
					
					strID = objRS("id_prodotto")								
					objProdotto.setIDProdotto(objRS("id_prodotto"))    
					objProdotto.setNomeProdotto(objRS("nome_prod"))
					objProdotto.setSommarioProdotto(objRS("sommario_prod"))
					objProdotto.setDescProdotto(objRS("desc_prod"))
					objProdotto.setPrezzo(objRS("prezzo"))
					objProdotto.setQtaDisp(objRS("qta_disp"))
					objProdotto.setAttivo(objRS("attivo"))
					objProdotto.setSconto(objRS("sconto"))
					objProdotto.setCodiceProd(objRS("codice_prod"))
					objProdotto.setIDTassaApplicata(objRS("id_tassa_applicata"))
					objProdotto.setProdType(objRS("prod_type"))
					objProdotto.setMaxDownload(objRS("max_download"))
					objProdotto.setMaxDownloadTime(objRS("max_download_time"))
					objProdotto.setTaxGroup(objRS("taxs_group"))		
					objProdotto.setMetaDescription(objRS("meta_description"))	
					objProdotto.setMetaKeyword(objRS("meta_keyword"))
					objProdotto.setPageTitle(objRS("page_title"))
					objProdotto.setEditBuyQta(objRS("edit_buy_qta"))

					objProdottoCache.add "id_prodotto", strID
					objProdottoCache.add "nome_prod", objProdotto.getNomeProdotto()
					objProdottoCache.add "sommario_prod", objProdotto.getSommarioProdotto()
					objProdottoCache.add "desc_prod", objProdotto.getDescProdotto()
					objProdottoCache.add "prezzo", objProdotto.getPrezzo()
					objProdottoCache.add "qta_disp", objProdotto.getQtaDisp()
					objProdottoCache.add "attivo", objProdotto.getAttivo()
					objProdottoCache.add "sconto", objProdotto.getSconto()
					objProdottoCache.add "codice_prod", objProdotto.getCodiceProd()
					objProdottoCache.add "id_tassa_applicata", objProdotto.getIDTassaApplicata()
					objProdottoCache.add "prod_type", objProdotto.getProdType()
					objProdottoCache.add "max_download", objProdotto.getMaxDownload()
					objProdottoCache.add "max_download_time", objProdotto.getMaxDownloadTime()
					objProdottoCache.add "taxs_group", objProdotto.getTaxGroup()
					objProdottoCache.add "meta_description", objProdotto.getMetaDescription()	
					objProdottoCache.add "meta_keyword", objProdotto.getMetaKeyword()		
					objProdottoCache.add "page_title", objProdotto.getPageTitle()	
					objProdottoCache.add "edit_buy_qta", objProdotto.getEditBuyQta()

					objDict.add strID, objProdotto
					objListaProdCache.add strID, objProdottoCache
					
					Set objProdottoCache = nothing
					Set objProdotto = nothing
					objRS.moveNext()
				loop
				
				if(CBool(bolAddTarget) OR CBool(bolAddAttachment))then
					Set objProdotto = new ProductsClass
					Set objFiles = new File4ProductsClass 

					for each j in objDict
						bolValid = true
						if(CBool(bolAddTarget))then				
							Set objListaTarget = objProdotto.getTargetPerProdotto(j)					
							if not(isEmpty(objListaTarget)) then
								objDict(j).setListaTarget(objListaTarget)
								
								Set objListaTargetCache = Server.CreateObject("Scripting.Dictionary")
								for each xt in objListaTarget
									Set objTargetCache = Server.CreateObject("Scripting.Dictionary")
									objTargetCache.add "id_target", xt
									objTargetCache.add "descrizione", objListaTarget(xt).getTargetDescrizione()
									objTargetCache.add "type", objListaTarget(xt).getTargetType()							
									objListaTargetCache.add xt, objTargetCache	
									'response.write("id_target:"&objTargetCache("id_target")&" - descrizione:"&objTargetCache("descrizione")&"<br>")						
									Set objTargetCache = nothing
								next	
								objListaProdCache(j).add "target_list", objListaTargetCache
								Set objListaTargetCache = nothing	
							else
								call objDict.remove(j)
								call objListaProdCache.remove(j)
								bolValid = false
							end if
							Set objListaTarget = nothing
						end if					
											
						if(bolValid AND CBool(bolAddAttachment)) then							
							on Error Resume Next				
							Set objListaFiles = objFiles.getFilePerProdotto(j)
							Set objListaFilesCache = Server.CreateObject("Scripting.Dictionary")
							if Err.number <> 0 then
								objListaFiles = null
							end if
						
							if not(isNull(objListaFiles)) then
								objProdotto.setFileXProdotto(objListaFiles)
								
								for each xf in objListaFiles
									Set objFilesCache = Server.CreateObject("Scripting.Dictionary")
									objFilesCache.add "id_attach", xf
									objFilesCache.add "id_prodotto", objListaFiles(xf).getProdottoID()
									objFilesCache.add "filename", objListaFiles(xf).getFileName()
									objFilesCache.add "content_type", objListaFiles(xf).getFileType()
									objFilesCache.add "path", objListaFiles(xf).getFilePath()
									objFilesCache.add "file_dida", objListaFiles(xf).getFileDida()
									objFilesCache.add "file_label", objListaFiles(xf).getFileTypeLabel()								
									objListaFilesCache.add xf, objFilesCache							
									Set objFilesCache = nothing
								next

								objListaProdCache(j).add "file_list", objListaFilesCache
								Set objListaFilesCache = nothing
							end if
							Set objListaFiles = nothing
						end if				
					next

					Set objFiles = Nothing
					Set objProdotto = nothing
				end if
				
				if (objDict.Count > 0) then
					Set findProdottiCached = objDict
					call ojbCache.store(cacheKey, objListaProdCache)
				else
					findProdottiCached = null			
				end if
					
				Set objListaProdCache = nothing					
				Set objDict = nothing    
			end if
			
			Set objListTarget = nothing
			Set objRS = Nothing
			Set objCommand = Nothing
			Set objDB = Nothing
			
			if Err.number <> 0 then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if 
		end if

		Set ojbCache = nothing		
	End Function
	
	Public Function findProdottoByID(id, attachments)
		on error resume next
		
		findProdottoByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM prodotti WHERE id_prodotto=?;"

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
			Dim objProdotto, objListaFiles
			Dim objFiles, objListaTarget
			
			Set objFiles = new File4ProductsClass	
			Set objProdotto = new ProductsClass	

			Dim this_id_prod
			this_id_prod = objRS("id_prodotto")
				
			objProdotto.setIDProdotto(this_id_prod)    
			objProdotto.setNomeProdotto(objRS("nome_prod"))
			objProdotto.setSommarioProdotto(objRS("sommario_prod"))
			objProdotto.setDescProdotto(objRS("desc_prod"))
			objProdotto.setPrezzo(objRS("prezzo"))
			objProdotto.setQtaDisp(objRS("qta_disp"))
			objProdotto.setAttivo(objRS("attivo"))
			objProdotto.setSconto(objRS("sconto"))
			objProdotto.setCodiceProd(objRS("codice_prod"))
			objProdotto.setIDTassaApplicata(objRS("id_tassa_applicata"))
			objProdotto.setProdType(objRS("prod_type"))
			objProdotto.setMaxDownload(objRS("max_download"))
			objProdotto.setMaxDownloadTime(objRS("max_download_time"))
			objProdotto.setTaxGroup(objRS("taxs_group"))		
			objProdotto.setMetaDescription(objRS("meta_description"))	
			objProdotto.setMetaKeyword(objRS("meta_keyword"))
			objProdotto.setPageTitle(objRS("page_title"))
			objProdotto.setEditBuyQta(objRS("edit_buy_qta"))
				
			bolValid = true

			Set objListaTarget = objProdotto.getTargetPerProdotto(this_id_prod)
			if not(isEmpty(objListaTarget)) then
				objProdotto.setListaTarget(objListaTarget)
				Set objListaTarget = nothing
			else
				Set objListaTarget = nothing
				findProdottoByID = null
				bolValid = false
			end if	
			
			if(bolValid)then			
				if(CBool(attachments)) then
					Set objListaFiles = objFiles.getFilePerProdotto(this_id_prod)
					if not(isEmpty(objListaFiles)) then
						objProdotto.setFileXProdotto(objListaFiles)
						Set objListaFiles = nothing
					else
						Set objListaFiles = nothing
					end if
				end if	
				
				Set findProdottoByID = objProdotto
			end if

			Set objFiles = Nothing
			Set objProdotto = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function findProdottoByParameter(paramType, paramValue, bolCache, attachments)		
		findProdottoByParameter = null
			
		Set objBase64 = new Base64Class

		if(bolCache) then
			'tento il recupero dell'oggetto dalla cache
			on error resume next
			Set ojbCache = new CacheClass
			cacheParamValue = objBase64.Base64Encode(paramValue)
			Set cachedObj = ojbCache.getItem("product-"&cacheParamValue)
					
			if (Instr(1, typename(cachedObj), "Dictionary", 1) > 0) then
				Set objProdottoC = new ProductsClass		
				objProdottoC.setIDProdotto(cachedObj("id_prodotto"))    
				objProdottoC.setNomeProdotto(cachedObj("nome_prod"))
				objProdottoC.setSommarioProdotto(cachedObj("sommario_prod"))
				objProdottoC.setDescProdotto(cachedObj("desc_prod"))
				objProdottoC.setPrezzo(cachedObj("prezzo"))
				objProdottoC.setQtaDisp(cachedObj("qta_disp"))
				objProdottoC.setAttivo(cachedObj("attivo"))
				objProdottoC.setSconto(cachedObj("sconto"))
				objProdottoC.setCodiceProd(cachedObj("codice_prod"))
				objProdottoC.setIDTassaApplicata(cachedObj("id_tassa_applicata"))
				objProdottoC.setProdType(cachedObj("prod_type"))
				objProdottoC.setMaxDownload(cachedObj("max_download"))
				objProdottoC.setMaxDownloadTime(cachedObj("max_download_time"))
				objProdottoC.setTaxGroup(cachedObj("taxs_group"))		
				objProdottoC.setMetaDescription(cachedObj("meta_description"))	
				objProdottoC.setMetaKeyword(cachedObj("meta_keyword"))
				objProdottoC.setPageTitle(cachedObj("page_title"))
				objProdottoC.setEditBuyQta(cachedObj("edit_buy_qta"))

				Set objListaTarget = Server.CreateObject("Scripting.Dictionary")
				Set objListaTargetTmp = cachedObj("target_list")
				for each xt in objListaTargetTmp
					Set objTarget = new Targetclass
					objTarget.setTargetID(xt)
					objTarget.setTargetDescrizione(objListaTargetTmp(xt)("descrizione"))
					objTarget.setTargetType(objListaTargetTmp(xt)("type"))	
					objListaTarget.add xt, objTarget
					Set objTarget = nothing		
				next
				Set objListaTargetTmp = nothing			
				objProdottoC.setListaTarget(objListaTarget)
				Set objListaTarget = nothing
				
				Set objListaFiles = Server.CreateObject("Scripting.Dictionary")
				Set objListaFilesTmp = cachedObj("file_list")
				for each xf in objListaFilesTmp
					Set objFiles = new File4ProductsClass
					objFiles.setFileID(xf)
					objFiles.setProdottoID(objListaFilesTmp(xf)("id_prodotto"))
					objFiles.setFileName(objListaFilesTmp(xf)("filename"))
					objFiles.setFileType(objListaFilesTmp(xf)("content_type"))
					objFiles.setFilePath(objListaFilesTmp(xf)("path"))
					objFiles.setFileDida(objListaFilesTmp(xf)("file_dida"))
					objFiles.setFileTypeLabel(objListaFilesTmp(xf)("file_label"))								
					objListaFiles.add xf, objFiles
					Set objFiles = nothing			
				next
				Set objListaFilesTmp = nothing
				if(objListaFiles.count>0)then
					objProdottoC.setFileXProdotto(objListaFiles)
				else
					objProdottoC.setFileXProdotto(null)
				end if
				Set objListaFiles = nothing
				
				Set findProdottoByParameter = objProdottoC
				Set objProdottoC = nothing			
			else
				findProdottoByParameter = null
			end if
			
			if Err.number <> 0 then
				findProdottoByParameter = null
				'response.write(Err.number&" - "&Err.description&"<br>")
			end if
		end if
		
		if not(Instr(1, typename(findProdottoByParameter), "ProductsClass", 1) > 0) then
			on error resume next
			Dim objDB, strSQL, objRS, objConn
			
			Select Case paramType
			Case "id"
				strSQL = "SELECT * FROM prodotti WHERE id_prodotto=?;"
			Case "name"
				strSQL = "SELECT * FROM prodotti WHERE nome_prod=?;"
			Case "code"
				strSQL = "SELECT * FROM prodotti WHERE codice_prod=?;"
			Case Else
				findProdottoByParameter = null
				Exit Function
			End Select

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()	
			Dim objCommand
			Set objCommand = Server.CreateObject("ADODB.Command")
			objCommand.ActiveConnection = objConn
			objCommand.CommandType=1
			objCommand.CommandText = strSQL
			
			Select Case paramType
			Case "id"
				objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,paramValue)
			Case "name"
				objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,paramValue)
			Case "code"
				objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,paramValue)
			Case Else
			End Select
			Set objRS = objCommand.Execute()		

			if not objRS.EOF then
				Dim objProdotto, objListaFiles
				Dim objFiles, objListaTarget
				
				Set objFiles = new File4ProductsClass	
				Set objProdotto = new ProductsClass
				Set objProdottoCache = Server.CreateObject("Scripting.Dictionary")		

				Dim this_id_prod
				this_id_prod = objRS("id_prodotto")
				
				objProdotto.setIDProdotto(this_id_prod)    
				objProdotto.setNomeProdotto(objRS("nome_prod"))
				objProdotto.setSommarioProdotto(objRS("sommario_prod"))
				objProdotto.setDescProdotto(objRS("desc_prod"))
				objProdotto.setPrezzo(objRS("prezzo"))
				objProdotto.setQtaDisp(objRS("qta_disp"))
				objProdotto.setAttivo(objRS("attivo"))
				objProdotto.setSconto(objRS("sconto"))
				objProdotto.setCodiceProd(objRS("codice_prod"))
				objProdotto.setIDTassaApplicata(objRS("id_tassa_applicata"))
				objProdotto.setProdType(objRS("prod_type"))
				objProdotto.setMaxDownload(objRS("max_download"))
				objProdotto.setMaxDownloadTime(objRS("max_download_time"))
				objProdotto.setTaxGroup(objRS("taxs_group"))		
				objProdotto.setMetaDescription(objRS("meta_description"))	
				objProdotto.setMetaKeyword(objRS("meta_keyword"))
				objProdotto.setPageTitle(objRS("page_title"))
				objProdotto.setEditBuyQta(objRS("edit_buy_qta"))

				if(bolCache) then
					objProdottoCache.add "id_prodotto", this_id_prod
					objProdottoCache.add "nome_prod", objProdotto.getNomeProdotto()
					objProdottoCache.add "sommario_prod", objProdotto.getSommarioProdotto()
					objProdottoCache.add "desc_prod", objProdotto.getDescProdotto()
					objProdottoCache.add "prezzo", objProdotto.getPrezzo()
					objProdottoCache.add "qta_disp", objProdotto.getQtaDisp()
					objProdottoCache.add "attivo", objProdotto.getAttivo()
					objProdottoCache.add "sconto", objProdotto.getSconto()
					objProdottoCache.add "codice_prod", objProdotto.getCodiceProd()
					objProdottoCache.add "id_tassa_applicata", objProdotto.getIDTassaApplicata()
					objProdottoCache.add "prod_type", objProdotto.getProdType()
					objProdottoCache.add "max_download", objProdotto.getMaxDownload()
					objProdottoCache.add "max_download_time", objProdotto.getMaxDownloadTime()
					objProdottoCache.add "taxs_group", objProdotto.getTaxGroup()
					objProdottoCache.add "meta_description", objProdotto.getMetaDescription()	
					objProdottoCache.add "meta_keyword", objProdotto.getMetaKeyword()		
					objProdottoCache.add "page_title", objProdotto.getPageTitle()	
					objProdottoCache.add "edit_buy_qta", objProdotto.getEditBuyQta()
				end if

				bolValid = true

				Set objListaTarget = objProdotto.getTargetPerProdotto(this_id_prod)
				if not(isEmpty(objListaTarget)) then
					objProdotto.setListaTarget(objListaTarget)
					
					if(bolCache) then
						Set objListaTargetCache = Server.CreateObject("Scripting.Dictionary")
						for each xt in objListaTarget
							Set objTargetCache = Server.CreateObject("Scripting.Dictionary")
							objTargetCache.add "id_target", xt
							objTargetCache.add "descrizione", objListaTarget(xt).getTargetDescrizione()
							objTargetCache.add "type", objListaTarget(xt).getTargetType()							
							objListaTargetCache.add xt, objTargetCache							
							Set objTargetCache = nothing
						next	
						objProdottoCache.add "target_list", objListaTargetCache
						Set objListaTargetCache = nothing	
					end if
					
					Set objListaTarget = nothing
				else
					Set objListaTarget = nothing
					findProdottoByParameter = null
					bolValid = false
				end if	
				
				if(bolValid)then
					if(bolCache) then
						Set objListaFilesCache = Server.CreateObject("Scripting.Dictionary")	
					end if					
					if(CBool(attachments)) then
						Set objListaFiles = objFiles.getFilePerProdotto(this_id_prod)
						if not(isEmpty(objListaFiles)) then
							objProdotto.setFileXProdotto(objListaFiles)
							
							if(bolCache) then							
								for each xf in objListaFiles
									Set objFilesCache = Server.CreateObject("Scripting.Dictionary")
									objFilesCache.add "id_attach", xf
									objFilesCache.add "id_prodotto", objListaFiles(xf).getProdottoID()
									objFilesCache.add "filename", objListaFiles(xf).getFileName()
									objFilesCache.add "content_type", objListaFiles(xf).getFileType()
									objFilesCache.add "path", objListaFiles(xf).getFilePath()
									objFilesCache.add "file_dida", objListaFiles(xf).getFileDida()
									objFilesCache.add "file_label", objListaFiles(xf).getFileTypeLabel()								
									objListaFilesCache.add xf, objFilesCache							
									Set objFilesCache = nothing
								next
							end if
							
							Set objListaFiles = nothing
						else
							Set objListaFiles = nothing
						end if
					end if
					
					if(bolCache) then					
						objProdottoCache.add "file_list", objListaFilesCache
						Set objListaFilesCache = nothing					
						call ojbCache.store("product-"&cacheParamValue, objProdottoCache)
					end if
					
					Set findProdottoByParameter = objProdotto		
				end if
				
				Set objProdottoCache = nothing
				Set objFiles = Nothing
				Set objProdotto = Nothing
			end if
			
			Set objRS = Nothing
			Set objCommand = Nothing
			Set objDB = Nothing
			
			if Err.number <> 0 then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if
		end if

		Set ojbCache = nothing
		Set objBase64 = nothing
	End Function

	Public Function getMaxIDProdotto()
		on error resume next
		
		getMaxIDProdotto = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT MAX(id_prodotto) AS id_prod FROM prodotti;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDProdotto = objRS("id_prod")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function countProdotti()
		on error resume next
		
		countProdotti = 0
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT count(*) AS counter FROM prodotti;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			countProdotti = objRS("counter")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function


	'*************************************************************************************	TARGET PER PRODOTTO	*************************************************************************************

	Public Sub insertTargetXProdotto(id_target, id_prod, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO target_x_prodotto(id_target, id_prodotto) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
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
	
	Public Sub insertTargetXProdottoNoTransaction(id_target, id_prod)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO target_x_prodotto(id_target, id_prodotto) VALUES("
		strSQL = strSQL & "?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteTargetXProdotto(id_prod, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM target_x_prodotto WHERE id_prodotto=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
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
		
	Public Sub deleteTargetXProdottoNoTransaction(id_prod)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM target_x_prodotto WHERE id_prodotto=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	

	Public Function getTargetPerProdotto(id_prod)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getTargetPerProdotto = null		
		strSQL = "SELECT target_x_prodotto.id_target, target.descrizione, target.type FROM target INNER JOIN target_x_prodotto ON target.id = target_x_prodotto.id_target WHERE target_x_prodotto.id_prodotto=? ORDER BY target_x_prodotto.id_target;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod)
		Set objRS = objCommand.Execute()			

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				Set objTarget = new Targetclass		
				strID = objRS("id_target")
				objTarget.setTargetID(strID)
				objTarget.setTargetDescrizione(objRS("descrizione"))
				objTarget.setTargetType(objRS("type"))	
				objDict.add strID, objTarget
				Set objTarget = nothing
				objRS.moveNext()
			loop
							
			Set getTargetPerProdotto = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	


	'*************************************************************************************	PRODOTTI CORRELATI	*************************************************************************************
	
	Public Sub insertRelationXProdotto(id_prod, id_prod_rel, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO relation_x_prodotto(id_prod, id_prod_rel) VALUES("
		strSQL = strSQL & "?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_rel)
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
	
	Public Sub insertRelationXProdottoNoTransaction(id_prod, id_prod_rel)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO relation_x_prodotto(id_prod, id_prod_rel) VALUES("
		strSQL = strSQL & "?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_rel)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteRelationXProdotto(id_prod_rel, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM relation_x_prodotto WHERE id_prod_rel=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_rel)
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
		
	Public Sub deleteRelationXProdottoNoTransaction(id_prod_rel)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM relation_x_prodotto WHERE id_prod_rel=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_rel)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
		
	Public Sub deleteAllRelationXProdotto(id_prod, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM relation_x_prodotto WHERE id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
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
		
	Public Sub deleteAllRelationXProdottoNoTransaction(id_prod)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM relation_x_prodotto WHERE id_prod=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	

	Public Function getRelationPerProdotto(id_prod)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getRelationPerProdotto = null		
		strSQL = "SELECT prodotti.*, relation_x_prodotto.id_prod_rel FROM prodotti INNER JOIN relation_x_prodotto ON prodotti.id_prodotto = relation_x_prodotto.id_prod_rel WHERE relation_x_prodotto.id_prod=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod)
		Set objRS = objCommand.Execute()

		if not(objRS.EOF) then		   
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
			Set objFiles = new File4ProductsClass  
			do while not objRS.EOF
				Set objProdotto = new ProductsClass
				strID = objRS("id_prodotto")
								
				objProdotto.setIDProdotto(strID)    
				objProdotto.setNomeProdotto(objRS("nome_prod"))
				objProdotto.setSommarioProdotto(objRS("sommario_prod"))
				objProdotto.setDescProdotto(objRS("desc_prod"))
				objProdotto.setPrezzo(objRS("prezzo"))
				objProdotto.setQtaDisp(objRS("qta_disp"))
				objProdotto.setAttivo(objRS("attivo"))
				objProdotto.setSconto(objRS("sconto"))
				objProdotto.setCodiceProd(objRS("codice_prod"))
				objProdotto.setIDTassaApplicata(objRS("id_tassa_applicata"))
				objProdotto.setProdType(objRS("prod_type"))
				objProdotto.setMaxDownload(objRS("max_download"))
				objProdotto.setMaxDownloadTime(objRS("max_download_time"))
				objProdotto.setTaxGroup(objRS("taxs_group"))		
				objProdotto.setMetaDescription(objRS("meta_description"))	
				objProdotto.setMetaKeyword(objRS("meta_keyword"))
				objProdotto.setPageTitle(objRS("page_title"))				
				objProdotto.setEditBuyQta(objRS("edit_buy_qta"))
				
				Set objListaTarget = objProdotto.getTargetPerProdotto(strID)
				if not(isEmpty(objListaTarget)) then
					objProdotto.setListaTarget(objListaTarget)
					Set objListaTarget = nothing
				else
					Set objListaTarget = nothing
					response.Redirect(Application("baseroot")&Application("error_page")&"?error=020")
				end if	
							
				Set objListaFiles = objFiles.getFilePerProdotto(strID)				
				if not(isEmpty(objListaFiles)) then
					objProdotto.setFileXProdotto(objListaFiles)
					Set objListaFiles = nothing
				else
					Set objListaFiles = nothing
				end if
									
				objDict.add strID, objProdotto
				Set objProdotto = nothing
				objRS.moveNext()
			loop
			
			Set objFiles = Nothing
			Set getRelationPerProdotto = objDict   
			Set objDict = nothing    
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function getListaProdotti4Relation()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objProdotto, objListaFiles
		Dim objFiles, objListaTarget
		getListaProdotti4Relation = null  
		'strSQL = "SELECT prodotti.id_prodotto, nome_prod, sommario_prod, desc_prod, prezzo, qta_disp, attivo, sconto, codice_prod, id_tassa_applicata, prod_type, max_download, max_download_time, taxs_group, categorie.descrizione FROM prodotti LEFT JOIN `target_x_prodotto` ON prodotti.id_prodotto = target_x_prodotto.id_prodotto LEFT JOIN target ON target.id = target_x_prodotto.id_target LEFT JOIN target_x_categoria ON target.id = target_x_categoria.id_target LEFT JOIN categorie ON target_x_categoria.id_categoria = categorie.id WHERE target.type=2 ORDER BY gerarchia;"
		strSQL = "SELECT prodotti.*, categorie.descrizione FROM prodotti LEFT JOIN `target_x_prodotto` ON prodotti.id_prodotto = target_x_prodotto.id_prodotto LEFT JOIN target ON target.id = target_x_prodotto.id_target LEFT JOIN target_x_categoria ON target.id = target_x_categoria.id_target LEFT JOIN categorie ON target_x_categoria.id_categoria = categorie.id WHERE target.type=2 ORDER BY gerarchia;"
		
		Set objProd = new ProductsClass
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()  
		Set objRS = objConn.Execute(strSQL)
		
		if not(objRS.EOF) then		   
			Set objDict = Server.CreateObject("Scripting.Dictionary") 
			do while not objRS.EOF
				Set objProdotto = new ProductsClass
				strID = objRS("id_prodotto")				
				objProdotto.setIDProdotto(strID)    
				objProdotto.setNomeProdotto(objRS("nome_prod"))
				objProdotto.setSommarioProdotto(objRS("sommario_prod"))
				objProdotto.setDescProdotto(objRS("desc_prod"))
				objProdotto.setPrezzo(objRS("prezzo"))
				objProdotto.setQtaDisp(objRS("qta_disp"))
				objProdotto.setAttivo(objRS("attivo"))
				objProdotto.setSconto(objRS("sconto"))
				objProdotto.setCodiceProd(objRS("codice_prod"))
				objProdotto.setIDTassaApplicata(objRS("id_tassa_applicata"))
				objProdotto.setProdType(objRS("prod_type"))
				objProdotto.setMaxDownload(objRS("max_download"))
				objProdotto.setMaxDownloadTime(objRS("max_download_time"))
				objProdotto.setTaxGroup(objRS("taxs_group"))
				objProdotto.setDescCatRelProd(objRS("descrizione"))		
				objProdotto.setMetaDescription(objRS("meta_description"))	
				objProdotto.setMetaKeyword(objRS("meta_keyword"))
				objProdotto.setPageTitle(objRS("page_title"))
				objProdotto.setEditBuyQta(objRS("edit_buy_qta"))
				
				'Set objListaTarget = objProdotto.getTargetPerProdotto(strID)
				'if not(isEmpty(objListaTarget)) then
					'objProdotto.setListaTarget(objListaTarget)
					'Set objListaTarget = nothing
				'else
					'Set objListaTarget = nothing
					'response.Redirect(Application("baseroot")&Application("error_page")&"?error=020")
				'end if
			
				objDict.add strID, objProdotto
				Set objProdotto = nothing
				objRS.moveNext()
			loop
			
			Set getListaProdotti4Relation = objDict   
			Set objDict = nothing    
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		
		
		Set objProd = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if  
	End Function

	Public Function getGerCatProd4Relation(id_prod)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		getGerCatProd4Relation = null
		strSQL = "SELECT DISTINCT(gerarchia) FROM categorie LEFT JOIN target_x_categoria ON categorie.id = target_x_categoria.id_categoria LEFT JOIN `target_x_prodotto` ON target_x_categoria.id_target = target_x_prodotto.id_target WHERE target_x_prodotto.id_prodotto=? LIMIT 1;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then	
			getGerCatProd4Relation = objRS("gerarchia")
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if 
	End Function


	'*************************************************************************************	MAIN FIELDS TRANSLATION	*************************************************************************************
	
	Public Sub insertFieldTranslation(id_prod, main_field ,lang_code, value, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO prodotto_main_field_translation(id_prod, main_field ,lang_code, value) VALUES("
		strSQL = strSQL & "?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,main_field)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,lang_code)
		if not(isNull(value)) then objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,value)
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
	
	Public Sub insertFieldTranslationNoTransaction(id_prod, main_field ,lang_code, value)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO prodotto_main_field_translation(id_prod, main_field ,lang_code, value) VALUES("
		strSQL = strSQL & "?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,main_field)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,lang_code)
		if not(isNull(value)) then objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,value)
		objCommand.Execute()		
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteFieldTranslation(id_prod, main_field ,lang_code, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM prodotto_main_field_translation WHERE id_prod=? AND main_field=? AND lang_code=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,main_field)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,lang_code)
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
		
	Public Sub deleteFieldTranslationNoTransaction(id_prod, main_field ,lang_code)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM prodotto_main_field_translation WHERE id_prod=? AND main_field=? AND lang_code=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,main_field)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,lang_code)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
	
	Public Function findFieldTranslation(main_field , lang_code, def)
		on error resume next
		
		findFieldTranslation = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM prodotto_main_field_translation WHERE id_prod=? AND main_field=? AND lang_code=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,Cint(getIDProdotto()))
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,main_field)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,2,lang_code)
		Set objRS = objCommand.Execute()		

		if not(objRS.EOF) then											
			findFieldTranslation = objRS("value")
			
			if(Trim(findFieldTranslation)="")then
				if(def=1)then
					Select Case main_field
					Case 1
						findFieldTranslation = getNomeProdotto()
					Case 2
						findFieldTranslation = getSommarioProdotto()
					Case 3
						findFieldTranslation =  getDescProdotto()	
					Case Else
						findFieldTranslation = null
					End Select
				else
					findFieldTranslation = null
				end if
			end if
		else
			if(def=1)then
				Select Case main_field
				Case 1
					findFieldTranslation = getNomeProdotto()
				Case 2
					findFieldTranslation = getSommarioProdotto()
				Case 3
					findFieldTranslation =  getDescProdotto()	
				Case Else
					findFieldTranslation = null
				End Select
			else
				findFieldTranslation = null
			end if
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	


	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = doubleValue
		
		'if (Application("dbType") = 0) then
			convertDoubleDelimiter = Replace(convertDoubleDelimiter, ".",",")
		'else		
			'convertDoubleDelimiter = Replace(convertDoubleDelimiter, ",",".")
		'end if			
	End Function

	'public Sub toString()
		'response.write ()
	'end Sub
End Class
%>