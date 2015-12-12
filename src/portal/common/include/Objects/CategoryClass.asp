<%
Class CategoryClass
	Private id_cat
	Private num_menu
	Private gerarchia_cat
	Private descrizione_cat
	Private type_cat
	Private id_template
	Private isVisible
	Private contiene_news
	Private contiene_prod
	Private meta_description
	Private meta_keyword
	Private page_title
	Private sub_domain_url
	
	Public Function getCatID()
		getCatID = id_cat
	End Function
	
	Public Sub setCatID(strID)
		id_cat = strID
	End Sub
	
	Public Function getNumMenu()
		getNumMenu = num_menu
	End Function
	
	Public Sub setNumMenu(strNumMenu)
		num_menu = strNumMenu
	End Sub
	
	Public Function getCatGerarchia()
		getCatGerarchia = gerarchia_cat
	End Function
	
	Public Sub setCatGerarchia(strGer)
		gerarchia_cat = strGer
	End Sub
	
	Public Function getCatDescrizione()
		getCatDescrizione = descrizione_cat
	End Function
	
	Public Sub setCatDescrizione(strDesc)
		descrizione_cat = strDesc
	End Sub
	
	Public Function getCatType()
		getCatType = type_cat
	End Function
	
	Public Sub setCatType(strType)
		type_cat = strType
	End Sub
	
	Public Function getIDTemplate()
		getIDTemplate = id_template
	End Function
	
	Public Sub setIDTemplate(numIDTemplate)
		id_template = numIDTemplate
	End Sub
	
	Public Function isCatVisible()
		isCatVisible = Cbool(isVisible)
	End Function
	
	Public Sub setCatVisible(bolIsVis)
		isVisible = bolIsVis
	End Sub
	
	Public Function contieneNews()
		contieneNews = Cbool(contiene_news)
	End Function
	
	Public Sub setContieneNews(bolContNews)
		contiene_news = bolContNews
	End Sub	
	
	Public Function contieneProd()
		contieneProd = Cbool(contiene_prod)
	End Function
	
	Public Sub setContieneProd(bolContProd)
		contiene_prod = bolContProd
	End Sub
	
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
	
	Public Function getSubDomainURL()
		getSubDomainURL = sub_domain_url
	End Function
	
	Public Sub setSubDomainURL(strSubDomainURL)
		sub_domain_url = strSubDomainURL
	End Sub	
	
	'*************************************************************************************************

	Public Function getMaxIDCat()
		on error resume next
		
		getMaxIDCat = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT MAX(categorie.id) as id FROM categorie;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			'objRS.MoveFirst()
			getMaxIDCat = objRS("id")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListaCategorie()
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objCategoria
		getListaCategorie = null		
		strSQL = "SELECT * FROM categorie ORDER BY gerarchia;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")			
			do while not objRS.EOF
				Set objCategoria = new CategoryClass
				strID = objRS("id")
				
				objCategoria.setCatID(objRS("id"))
				objCategoria.setNumMenu(objRS("num_menu"))			
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(objRS("contiene_news"))
				objCategoria.setContieneProd(objRS("contiene_prod"))
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))	
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))	
				objCategoria.setPageTitle(objRS("page_title"))		
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))			
				
				objDict.add strID, objCategoria
				Set objCategoria = nothing
				objRS.moveNext()
			loop
							
			Set getListaCategorie = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function getListaCategorieByMenu(iNumMenu)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict, objCategoria
		getListaCategorieByMenu = null		
		strSQL = "SELECT * FROM categorie WHERE num_menu = ? ORDER BY gerarchia;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,iNumMenu)
		Set objRS = objCommand.Execute()

		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")			
			do while not objRS.EOF
				Set objCategoria = new CategoryClass
				strID = objRS("id")
				
				objCategoria.setCatID(objRS("id"))	
				objCategoria.setNumMenu(objRS("num_menu"))			
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(objRS("contiene_news"))
				objCategoria.setContieneProd(objRS("contiene_prod"))
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))	
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))	
				objCategoria.setPageTitle(objRS("page_title"))	
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))								
				
				objDict.add strID, objCategoria
				Set objCategoria = nothing
				objRS.moveNext()
			loop
							
			Set getListaCategorieByMenu = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function findCategoriaByID(id)
		on error resume next
		
		findCategoriaByID = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE id=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else
			Dim objCategoria
			Set objCategoria = new CategoryClass
			
			objCategoria.setCatID(objRS("id"))	
			objCategoria.setNumMenu(objRS("num_menu"))			
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))				
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))	
			objCategoria.setPageTitle(objRS("page_title"))	
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))		

			Set findCategoriaByID = objCategoria
			Set objCategoria = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
		
	Public Function findFirstAvailableCategoria()
		on error resume next
		
		findFirstAvailableCategoria = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie ORDER BY gerarchia;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then
			Dim objCategoria
			Set objCategoria = new CategoryClass
			
			objCategoria.setCatID(objRS("id"))	
			objCategoria.setNumMenu(objRS("num_menu"))			
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))				
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))	
			objCategoria.setPageTitle(objRS("page_title"))	
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))		

			Set findFirstAvailableCategoria = objCategoria
			Set objCategoria = Nothing
		end if		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

		
	Public Function findCategoriaByGerarchia(iGerarchia)
		on error resume next
		
		findCategoriaByGerarchia = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE gerarchia=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else
			Dim objCategoria
			Set objCategoria = new CategoryClass
			
			objCategoria.setCatID(objRS("id"))	
			objCategoria.setNumMenu(objRS("num_menu"))			
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))		
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))
			objCategoria.setPageTitle(objRS("page_title"))	
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))					

			Set findCategoriaByGerarchia = objCategoria
			Set objCategoria = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

		
	Public Function findExsitingCategoriaByGerarchia(iGerarchia)
		on error resume next
		
		findExsitingCategoriaByGerarchia = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE gerarchia=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else
			Dim objCategoria
			Set objCategoria = new CategoryClass
			
			objCategoria.setCatID(objRS("id"))	
			objCategoria.setNumMenu(objRS("num_menu"))			
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))		
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))	
			objCategoria.setPageTitle(objRS("page_title"))
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))					

			Set findExsitingCategoriaByGerarchia = objCategoria
			Set objCategoria = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

		
	Public Function findCategoriaByDescription(strDescrizione)
		on error resume next
		
		findCategoriaByDescription = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE descrizione=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		Set objRS = objCommand.Execute()

		if objRS.EOF then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else
			Dim objCategoria
			Set objCategoria = new CategoryClass
			
			objCategoria.setCatID(objRS("id"))	
			objCategoria.setNumMenu(objRS("num_menu"))			
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))				
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))	
			objCategoria.setPageTitle(objRS("page_title"))	
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))		

			Set findCategoriaByDescription = objCategoria
			Set objCategoria = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function findExsitingGerarchia(iGerarchia)
		on error resume next
		
		findExsitingGerarchia = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE gerarchia=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia)
		Set objRS = objCommand.Execute()	

		if objRS.EOF then
			findExsitingGerarchia = null	
		else			
			findExsitingGerarchia = objRS("id")
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function findFirstCategoriaWithNews()
		on error resume next
		
		findFirstCategoriaWithNews = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE contiene_news=1 ORDER BY gerarchia;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then
			Dim objCategoria
			Set objCategoria = new CategoryClass
			
			objCategoria.setCatID(objRS("id"))	
			objCategoria.setNumMenu(objRS("num_menu"))			
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))				
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))	
			objCategoria.setPageTitle(objRS("page_title"))	
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))		

			Set findFirstCategoriaWithNews = objCategoria
			Set objCategoria = Nothing
		end if		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function findFirstSubCategoriaWithNews(iGerarchia)
		on error resume next
		
		findFirstSubCategoriaWithNews = null
				
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE gerarchia LIKE ? ORDER BY gerarchia;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia&".%")
		Set objRS = objCommand.Execute()		
		
		Dim hasNews, continue
		continue = true
		do while (not(objRS.EOF) AND continue)
			hasNews = objRS("contiene_news")
			
			if(hasNews) then				
				Set objCategoria = new CategoryClass				
				objCategoria.setCatID(objRS("id"))
				objCategoria.setNumMenu(objRS("num_menu"))				
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(hasNews)
				objCategoria.setContieneProd(objRS("contiene_prod"))
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))		
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))	
				objCategoria.setPageTitle(objRS("page_title"))	
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))
				Set findFirstSubCategoriaWithNews = objCategoria				
				Set objCategoria = nothing
				continue = false
			end if
			objRS.moveNext()
		loop	
			
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function findFirstCategoriaWithProd()
		on error resume next
		
		findFirstCategoriaWithProd = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE contiene_prod=1 ORDER BY gerarchia;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if not(objRS.EOF) then
			Dim objCategoria
			Set objCategoria = new CategoryClass
			
			objCategoria.setCatID(objRS("id"))	
			objCategoria.setNumMenu(objRS("num_menu"))			
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))				
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))	
			objCategoria.setPageTitle(objRS("page_title"))	
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))		

			Set findFirstCategoriaWithProd = objCategoria
			Set objCategoria = Nothing
		end if		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function		

		
	Public Function findFirstSubCategoriaWithProd(iGerarchia)
		on error resume next
		
		findFirstSubCategoriaWithProd = null
				
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE gerarchia LIKE ? ORDER BY gerarchia;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia&".%")
		Set objRS = objCommand.Execute()	
		
		Dim hasProd, continue
		continue = true
		do while (not(objRS.EOF) AND continue)
			hasProd = objRS("contiene_prod")
			
			if(hasProd) then				
				Set objCategoria = new CategoryClass				
				objCategoria.setCatID(objRS("id"))
				objCategoria.setNumMenu(objRS("num_menu"))				
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(objRS("contiene_news"))
				objCategoria.setContieneProd(hasProd)
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))		
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))	
				objCategoria.setPageTitle(objRS("page_title"))	
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))
				Set findFirstSubCategoriaWithProd = objCategoria				
				Set objCategoria = nothing				
				continue = false
			end if
			objRS.moveNext()
		loop	
			
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

		
	Public Function findChildCategoriaStartWithGerarchia(iGerarchia)
		on error resume next
		
		findChildCategoriaStartWithGerarchia = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE gerarchia LIKE ? ORDER BY gerarchia;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia&".%")
		Set objRS = objCommand.Execute()	

		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else
			Set objDict = Server.CreateObject("Scripting.Dictionary")			
			Dim objCategoria
			do while not objRS.EOF
				Set objCategoria = new CategoryClass
				strID = objRS("id")
				
				objCategoria.setCatID(objRS("id"))
				objCategoria.setNumMenu(objRS("num_menu"))				
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(objRS("contiene_news"))
				objCategoria.setContieneProd(objRS("contiene_prod"))
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))		
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))	
				objCategoria.setPageTitle(objRS("page_title"))	
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))					
				
				objDict.add strID, objCategoria
				Set objCategoria = nothing
				objRS.moveNext()
			loop	
							
			Set findChildCategoriaStartWithGerarchia = objDict		
			Set objDict = nothing		
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

		
	Public Function findExsitingChildCategoriaStartWithGerarchia(iGerarchia)
		on error resume next
		
		findExsitingChildCategoriaStartWithGerarchia = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE gerarchia LIKE ? ORDER BY gerarchia LIMIT 1;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia&".%")
		Set objRS = objCommand.Execute()	

		if objRS.EOF then
		else	
			Dim objCategoria
			Set objCategoria = new CategoryClass
			strID = objRS("id")			
			objCategoria.setCatID(objRS("id"))
			objCategoria.setNumMenu(objRS("num_menu"))				
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))		
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))	
			objCategoria.setPageTitle(objRS("page_title"))	
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))					
			
			Set findExsitingChildCategoriaStartWithGerarchia = objCategoria
			Set objCategoria = nothing	
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
		
	Public Function findCategoriaStartWithGerarchia(iGerarchia)
		on error resume next
		
		findCategoriaStartWithGerarchia = null
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT * FROM categorie WHERE gerarchia LIKE ? ORDER BY gerarchia;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia&"%")
		Set objRS = objCommand.Execute()		

		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else
			Set objDict = Server.CreateObject("Scripting.Dictionary")			
			Dim objCategoria
			do while not objRS.EOF
				Set objCategoria = new CategoryClass
				strID = objRS("id")
				
				objCategoria.setCatID(objRS("id"))	
				objCategoria.setNumMenu(objRS("num_menu"))			
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(objRS("contiene_news"))
				objCategoria.setContieneProd(objRS("contiene_prod"))
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))		
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))
				objCategoria.setPageTitle(objRS("page_title"))			
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))						
				
				objDict.add strID, objCategoria
				Set objCategoria = nothing
				objRS.moveNext()
			loop	
			Set objCategoria = Nothing
							
			Set findCategoriaStartWithGerarchia = objDict		
			Set objDict = nothing	
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

		
	Public Function findFirstCategoriaByTargetID(id_target)
		on error resume next
		
		findFirstCategoriaByTargetID = null
		
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "SELECT * FROM categorie WHERE id IN (SELECT target_x_categoria.id_categoria FROM target_x_categoria WHERE target_x_categoria.id_target=?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_target)
		Set objRS = objCommand.Execute()				

		if objRS.EOF then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else			
			Dim objCategoria
			Set objCategoria = new CategoryClass
			
			objCategoria.setCatID(objRS("id"))
			objCategoria.setNumMenu(objRS("num_menu"))				
			objCategoria.setCatGerarchia(objRS("gerarchia"))
			objCategoria.setCatDescrizione(objRS("descrizione"))
			objCategoria.setCatType(objRS("type"))
			objCategoria.setContieneNews(objRS("contiene_news"))
			objCategoria.setContieneProd(objRS("contiene_prod"))
			objCategoria.setCatVisible(objRS("visibile"))
			objCategoria.setIDTemplate(objRS("id_template"))			
			objCategoria.setMetaDescription(objRS("meta_description"))	
			objCategoria.setMetaKeyword(objRS("meta_keyword"))	
			objCategoria.setPageTitle(objRS("page_title"))		
			objCategoria.setSubDomainURL(objRS("sub_domain_url"))		

			Set findFirstCategoriaByTargetID = objCategoria
			Set objCategoria = Nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

		
	Public Function findCategorieByTargetID(id_target)
		on error resume next
		
		findCategorieByTargetID = null
		
		Dim objDB, strSQL, objRS, objConn		
		strSQL = "SELECT * FROM categorie WHERE id IN(SELECT target_x_categoria.id_categoria FROM target_x_categoria WHERE target_x_categoria.id_target=?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_target)
		Set objRS = objCommand.Execute()	

		if objRS.EOF then
		else
			Set findCategorieByTargetID = Server.CreateObject("Scripting.Dictionary")
			do while not objRS.EOF						
				Dim objCategoria, idCat
				Set objCategoria = new CategoryClass
				idCat = objRS("id") 
				objCategoria.setCatID(idCat)
				objCategoria.setNumMenu(objRS("num_menu"))				
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(objRS("contiene_news"))
				objCategoria.setContieneProd(objRS("contiene_prod"))
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))				
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))
				objCategoria.setPageTitle(objRS("page_title"))
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))				
	
				findCategorieByTargetID.add idCat, objCategoria
				Set objCategoria = Nothing
				objRS.moveNext()
			loop			
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

		
	Public Function findCategorieByType(catType)
		on error resume next
		
		findCategorieByType = null
		
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "SELECT * FROM categorie WHERE type=? ORDER BY gerarchia;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,catType)
		Set objRS = objCommand.Execute()	

		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else
			Set findCategorieByType = Server.CreateObject("Scripting.Dictionary")
			do while not objRS.EOF		
				
				Dim objCategoria, idCat
				Set objCategoria = new CategoryClass
				idCat = objRS("id")
				objCategoria.setCatID(idCat)	
				objCategoria.setNumMenu(objRS("num_menu"))			
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(objRS("contiene_news"))
				objCategoria.setContieneProd(objRS("contiene_prod"))
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))		
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))
				objCategoria.setPageTitle(objRS("page_title"))
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))						
	
				findCategorieByType.add idCat, objCategoria
				Set objCategoria = Nothing
				objRS.moveNext()
			loop			
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

		
	Public Function findCategorieByTypeAndMixed(catType)
		on error resume next
		
		findCategorieByTypeAndMixed = null
		
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "SELECT * FROM categorie WHERE type=? OR type=? ORDER BY gerarchia;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,catType)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,Application("strMixedCat"))
		Set objRS = objCommand.Execute()	

		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=013")		
		else
			Set findCategorieByTypeAndMixed = Server.CreateObject("Scripting.Dictionary")
			do while not objRS.EOF		
				
				Dim objCategoria, idCat
				Set objCategoria = new CategoryClass
				idCat = objRS("id")
				objCategoria.setCatID(idCat)	
				objCategoria.setNumMenu(objRS("num_menu"))			
				objCategoria.setCatGerarchia(objRS("gerarchia"))
				objCategoria.setCatDescrizione(objRS("descrizione"))
				objCategoria.setCatType(objRS("type"))
				objCategoria.setContieneNews(objRS("contiene_news"))
				objCategoria.setContieneProd(objRS("contiene_prod"))
				objCategoria.setCatVisible(objRS("visibile"))
				objCategoria.setIDTemplate(objRS("id_template"))		
				objCategoria.setMetaDescription(objRS("meta_description"))	
				objCategoria.setMetaKeyword(objRS("meta_keyword"))
				objCategoria.setPageTitle(objRS("page_title"))
				objCategoria.setSubDomainURL(objRS("sub_domain_url"))						
	
				findCategorieByTypeAndMixed.add idCat, objCategoria
				Set objCategoria = Nothing
				objRS.moveNext()
			loop			
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Function checkEmptyCategory(objCategoriaCheck, checkDeep)
		checkEmptyCategory = null
		if not(objCategoriaCheck.contieneNews()) AND not(objCategoriaCheck.contieneProd()) then
			if(checkDeep)then
				On Error Resume Next
				Set checkEmptyCategory = objCategoriaCheck.findFirstSubCategoriaWithNews(objCategoriaCheck.getCatGerarchia())
				
				if(isNull(checkEmptyCategory)) then
					Set checkEmptyCategory = objCategoriaCheck.findFirstSubCategoriaWithProd(objCategoriaCheck.getCatGerarchia())
				end if
				
				if(Err.number<>0)then
					checkEmptyCategory = null
				end if
			end if
		else
			Set checkEmptyCategory = objCategoriaCheck
		end if
	End Function
	
				
	Public Function insertCategoria(iNumMenu, iGerarchia, strDescrizione, strTypeCat, bolContNews, bolContProd, bolVisible, numIDTemplate, strMetaDesc, strMetaKey, strPageTitle, strSubDomainURL, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		insertCategoria = -1
		
		strSQL = "INSERT INTO categorie(num_menu, gerarchia, descrizione, type, contiene_news, contiene_prod, visibile, id_template, meta_description, meta_keyword, page_title, sub_domain_url) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,?,?,?,?);"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,iNumMenu)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTypeCat)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolContNews)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolContProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolVisible)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strSubDomainURL)
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(categorie.id) as id FROM categorie;")
		if not (objRS.EOF) then
			'objRS.MoveFirst()
			insertCategoria = objRS("id")	
		end if		
		Set objRS = Nothing		
		Set objCommand = Nothing
		
		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
	
				
	Public Function insertCategoriaNoTransaction(iNumMenu, iGerarchia, strDescrizione, strTypeCat, bolContNews, bolContProd, bolVisible, numIDTemplate, strMetaDesc, strMetaKey, strPageTitle, strSubDomainURL)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		insertCategoriaNoTransaction = -1
		
		strSQL = "INSERT INTO categorie(num_menu, gerarchia, descrizione, type, contiene_news, contiene_prod, visibile, id_template, meta_description, meta_keyword, page_title, sub_domain_url) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?,?,?,?,?,?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,iNumMenu)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTypeCat)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolContNews)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolContProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolVisible)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strSubDomainURL)
		objCommand.Execute()

		Set objRS = objConn.Execute("SELECT max(categorie.id) as id FROM categorie;")
		if not (objRS.EOF) then
			'objRS.MoveFirst()
			insertCategoriaNoTransaction = objRS("id")	
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyCategoria(id, iNumMenu, iGerarchia, strDescrizione, strTypeCat, bolContNews, bolContProd, bolVisible, numIDTemplate, strMetaDesc, strMetaKey, strPageTitle, strSubDomainURL, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE categorie SET "
		strSQL = strSQL & "num_menu=?,"
		strSQL = strSQL & "gerarchia=?,"
		strSQL = strSQL & "descrizione=?,"
		strSQL = strSQL & "type=?,"
		strSQL = strSQL & "contiene_news=?,"
		strSQL = strSQL & "contiene_prod=?,"
		strSQL = strSQL & "visibile=?,"
		strSQL = strSQL & "id_template=?,"
		strSQL = strSQL & "meta_description=?,"
		strSQL = strSQL & "meta_keyword=?,"
		strSQL = strSQL & "page_title=?,"
		strSQL = strSQL & "sub_domain_url=?"
		strSQL = strSQL & " WHERE id=?;" 

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,iNumMenu)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTypeCat)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolContNews)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolContProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolVisible)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strSubDomainURL)
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
		
	Public Sub modifyCategoriaNoTransaction(id, iNumMenu, iGerarchia, strDescrizione, strTypeCat, bolContNews, bolContProd, bolVisible, numIDTemplate, strMetaDesc, strMetaKey, strPageTitle, strSubDomainURL)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "UPDATE categorie SET "
		strSQL = strSQL & "num_menu=?,"
		strSQL = strSQL & "gerarchia=?,"
		strSQL = strSQL & "descrizione=?,"
		strSQL = strSQL & "type=?,"
		strSQL = strSQL & "contiene_news=?,"
		strSQL = strSQL & "contiene_prod=?,"
		strSQL = strSQL & "visibile=?,"
		strSQL = strSQL & "id_template=?,"
		strSQL = strSQL & "meta_description=?,"
		strSQL = strSQL & "meta_keyword=?,"
		strSQL = strSQL & "page_title=?,"
		strSQL = strSQL & "sub_domain_url=?"
		strSQL = strSQL & " WHERE id=?;" 
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,iNumMenu)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,iGerarchia)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTypeCat)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolContNews)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolContProd)
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,bolVisible)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,numIDTemplate)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaDesc)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strMetaKey)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,strPageTitle)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,strSubDomainURL)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteCategoria(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, strSQLDelTarget
		strSQLDelTarget = "DELETE FROM target_x_categoria WHERE id_categoria=?;"
		strSQL = "DELETE FROM categorie WHERE id=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	

		Dim objCommand, objCommand2
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQLDelTarget
		objCommand2.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)

		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand.Execute()
		end if				

		objCommand2.Execute()

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
	
	
	Public Sub changeCategoryVisibility(id_category, stato_cat)
		on error resume next
		
		Dim objDB, strSQL, objRS
		Dim objConn		
		
		strSQL = "UPDATE categorie SET "
		strSQL = strSQL & "visibile=?"	
		strSQL = strSQL & " WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,11,1,,stato_cat)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_category)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Sub	


	Public Function findCategoriaAssociations(id_categoria)
		on error resume next
		Dim objDB, strSQL, strSQL2, strSQL3, objRS, objConn
		findCategoriaAssociations = false		
		strSQL = "SELECT target_x_categoria.id_categoria FROM target_x_categoria WHERE target_x_categoria.id_categoria=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_categoria)
		Set objRS = objCommand.Execute()	

		if not(objRS.EOF) then							
			findCategoriaAssociations = true				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function findGerarchiaByTemplateDirectory(dir_template)	
		on error resume next
		Dim objDB, strSQL, strSQL2, strSQL3, objRS, objConn
		findGerarchiaByTemplateDirectory = null	
		strSQL = "SELECT gerarchia FROM `categorie` WHERE id_template IN(SELECT id FROM `template_disponibili` WHERE dir_template=?);"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,50,Trim(dir_template))
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then							
			findGerarchiaByTemplateDirectory = objRS("gerarchia")				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			findGerarchiaByTemplateDirectory = null
		end if		
	End Function
	
	Public Sub insertTargetXCategoria(id_target, id_cat, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO target_x_categoria(id_target, id_categoria) VALUES("
		strSQL = strSQL & "?,?);"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_cat)
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
		
	Public Sub deleteTargetXCategoria(id_cat, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM target_x_categoria WHERE id_categoria=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_cat)
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
	
	Public Sub insertTargetXCategoriaNoTransaction(id_target, id_cat)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		
		strSQL = "INSERT INTO target_x_categoria(id_target, id_categoria) VALUES("
		strSQL = strSQL & "?,?);"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_target)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_cat)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteTargetXCategoriaNoTransaction(id_cat)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
		strSQL = "DELETE FROM target_x_categoria WHERE id_categoria=?;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_cat)
		objCommand.Execute()	
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	

	Public Function getTargetPerCategoria(id_cat)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getTargetPerCategoria = null		
		strSQL = "SELECT target_x_categoria.id_target, target.descrizione, target.type FROM target INNER JOIN target_x_categoria ON target.id = target_x_categoria.id_target WHERE target_x_categoria.id_categoria=? ORDER BY target.type, target.descrizione;"
	
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_cat)
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
							
			Set getTargetPerCategoria = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	
	'**************************** FUNZIONI DI SUPPORTO PER GESTIRE I TEMPLATE PER LINGUA *********
	Public Function findLangTemplateXCategoria(lang_code, do_cascade)
		on error resume next
		Dim objDB, strSQL, strSQL2, strSQL3, objRS, objConn
		findLangTemplateXCategoria = null		
		strSQL = "SELECT id_template FROM template_x_categoria WHERE id_categoria=? AND lang_code=?;"
				
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_cat)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,lang_code)
		Set objRS = objCommand.Execute()	

		if not(objRS.EOF) then							
			findLangTemplateXCategoria = objRS("id_template")
		else
			if(do_cascade)then
				findLangTemplateXCategoria = id_template
			end if
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Sub insertLangTemplateXCategoria(id_cat, id_template, lang_code, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO template_x_categoria(id_categoria, id_template, lang_code) VALUES("
		strSQL = strSQL & "?,?,?);"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_cat)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_template)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,lang_code)
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
		
	Public Sub deleteLangTemplateXCategoria(id_cat, id_template, lang_code, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM template_x_categoria WHERE id_categoria=? AND id_template=? AND lang_code=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_cat)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_template)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,10,lang_code)
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
		
	Public Sub deleteAllLangTemplateXCategoria(id_cat, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "DELETE FROM template_x_categoria WHERE id_categoria=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_cat)
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
	
	
	'**************************** FUNZIONE DI SUPPORTO PER L'ORDINAMENTO DI SCRIPTING DICTIONARY *********

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