<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->

<%
if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if
	
	Dim id_categoria, iNumMenu, iGerarchia, strDescrizione, catType, bolContNews, bolContProd, bolVisible
	Dim reqTargets, arrTarget, idTemplate, metaDescription, metaKeyword, pageTitle, sub_domain_url
	Dim new_target, set_target_to_users, tmp_new_target, old_target, insert_new_target, set_target_to_categoria
	Dim catIdTmp, catConflicts, objLangTmp, objLangList
	id_categoria = request("id_categoria")
	iNumMenu = request("num_menu")
	iGerarchia = request("gerarchia")
	strDescrizione = request("descrizione")
	catType = request("cat_type")
	bolContNews = request("contiene_news")
	bolContProd = request("contiene_prod")
	bolVisible = request("visibile")
	idTemplate = request("id_template")
	new_target = request("new_target")
	old_target = request("old_target")
	set_target_to_users = request("set_target_to_users")
	set_target_to_categoria = request("set_target_to_categoria")
	insert_new_target = request("insert_new_target")
	metaDescription = request("meta_description")
	metaKeyword = request("meta_keyword")
	pageTitle = request("page_title") 
	sub_domain_url = request("sub_domain_url")
	
	reqTargets = request("ListTarget")
	arrTarget = split(reqTargets, "|", -1, 1)
	
	Dim objCategoria, objTarget, bolDelCategoria
	Set objCategoria = New CategoryClass
	Set objTarget = New TargetClass
	bolDelCategoria = request("delete_categoria")
	
	Dim objLogger, targetType
	Set objLogger = New LogClass
	
	if (Cint(id_categoria) <> -1) then
		
		if(strComp(bolDelCategoria, "del", 1) = 0) then
			if(objCategoria.findCategoriaAssociations(id_categoria)) then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=014")		
			else
				call objCategoria.deleteCategoria(id_categoria)
				call objLogger.write("cancellata categoria --> id: "&id_categoria&"; gerarchia: "&iGerarchia&"; descrizione: "&strDescrizione, objUserLogged.getUserName(), "info")
				
				Set objTarget = nothing
				response.Redirect(Application("baseroot")&"/editor/categorie/ListaCategorie.asp")	
			end if
		
		end if
		
		'**** verifico che non esista già una categoria con la stessa gerarchia
		catConflicts = false
		catIdTmp = objCategoria.findExsitingGerarchia(iGerarchia)
		if not(isNull(catTmp)) then
			if not(strComp(catIdTmp, id_categoria, 1) = 0) then
				catConflicts = true
			end if
		end if
		
		'**** verifico che non esista già una categoria con la stessa descrizione
		On Error Resume Next
		Set catIdTmp = objCategoria.findCategoriaByDescription(strDescrizione)
		if (strComp(typename(catIdTmp), "CategoryClass", 1) = 0) then
			if not(strComp(catIdTmp.getCatID(), id_categoria, 1) = 0) then
				catConflicts = true
			end if
		end if
		Set catIdTmp = nothing
		if(Err.number <>0)then
			catConflicts = false
		end if
		
		if not(catConflicts = true) then 		
			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans			
			
			call objCategoria.modifyCategoria(id_categoria, iNumMenu, iGerarchia, strDescrizione, catType, bolContNews, bolContProd, bolVisible, idTemplate, metaDescription, metaKeyword, pageTitle, sub_domain_url ,objConn)
			call objLogger.write("modificata categoria --> id: "&id_categoria&"; gerarchia: "&iGerarchia&"; descrizione: "&strDescrizione, objUserLogged.getUserName(), "info")
				
			call objCategoria.deleteTargetXCategoria(id_categoria, objConn)
					
			if not(isNull(arrTarget)) then
				for each x in arrTarget
					call objCategoria.insertTargetXCategoria(x, id_categoria, objConn)
				next
			else
				objConn.RollBackTrans
				Set objDB = nothing
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=005")
			end if
			
			'**** cancello e inserisco i template per lingua
			call objCategoria.deleteAllLangTemplateXCategoria(id_categoria, objConn)

			if (Cint(idTemplate) <> -1) then
				'*** RECUPERO LA LISTA DI LANGUAGE DISPONIBILI
				Set objLangTmp = New LanguageClass
				Set objLangList = objLangTmp.getListaLanguage()
				Set objLangTmp = nothing
					
				For Each x In objLangList
					lang_code_cat = UCase(objLangList(x).getLanguageDescrizione())
					id_lang_template = request("id_template_"&lang_code_cat)
					if (Cint(id_lang_template) <> -1) then
						call objCategoria.insertLangTemplateXCategoria(id_categoria, id_lang_template, lang_code_cat, objConn)
					end if
				Next		
				Set objLangList = nothing
			end if
					
			Set objCategoria = nothing
			Set catTmp = nothing
							
			if(catType = Application("strContentCat")) then
				targetType = 1
			elseif(catType = Application("strProdCat")) then
				targetType = 2
			else
				targetType = 4
			end if
			
			if(insert_new_target = 1) then
				if not(isNull(objTarget.findTargetByDescEq(old_target, objConn))) then						
					select case targetType		
					case 1	
						if (strComp(typename(objTarget.findTargetByDescEqAndType(old_target, 1, objConn)), "TargetClass", 1) = 0) then
							Set objTmpTarget = objTarget.findTargetByDescEqAndType(old_target, 1, objConn)
							call objTarget.modifyTarget(objTmpTarget.getTargetID(), new_target, targetType, 0, 0, objConn)
							Set objTmpTarget = nothing		
						else
							call objTarget.insertTarget(new_target, targetType, 0, 0, objConn)
						end if
					case 2	
						if (strComp(typename(objTarget.findTargetByDescEqAndType(old_target, 2, objConn)), "TargetClass", 1) = 0) then
							Set objTmpTarget = objTarget.findTargetByDescEqAndType(old_target, 2, objConn)
							call objTarget.modifyTarget(objTmpTarget.getTargetID(), new_target, targetType, 0, 0, objConn)
							Set objTmpTarget = nothing		
						else
							call objTarget.insertTarget(new_target, targetType, 0, 0, objConn)								
						end if			
					case 4
						if (strComp(typename(objTarget.findTargetByDescEqAndType(old_target, 1, objConn)), "TargetClass", 1) = 0) then
							Set objTmpTarget = objTarget.findTargetByDescEqAndType(old_target, 1, objConn)
							call objTarget.modifyTarget(objTmpTarget.getTargetID(), new_target, 1, 0, 0, objConn)
							Set objTmpTarget = nothing		
						else
							call objTarget.insertTarget(new_target, 1, 0, 0, objConn)							
						end if
						if (strComp(typename(objTarget.findTargetByDescEqAndType(old_target, 2, objConn)), "TargetClass", 1) = 0) then
							Set objTmpTarget = objTarget.findTargetByDescEqAndType(old_target, 2, objConn)
							call objTarget.modifyTarget(objTmpTarget.getTargetID(), new_target, 2, 0, 0, objConn)
							Set objTmpTarget = nothing		
						else
							call objTarget.insertTarget(new_target, 2, 0, 0, objConn)									
						end if					
					case else
					end select
				else
					select case targetType		
					case 1	
						call objTarget.insertTarget(new_target, targetType, 0, 0, objConn)
					case 2	
						call objTarget.insertTarget(new_target, targetType, 0, 0, objConn)			
					case 4
						call objTarget.insertTarget(new_target, 1, 0, 0, objConn)
						call objTarget.insertTarget(new_target, 2, 0, 0, objConn)				
					case else
					end select
				end if
			end if
						
			if objConn.Errors.Count = 0 then
				objConn.CommitTrans
			else
				objConn.RollBackTrans
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if			
			Set objDB = nothing
						
			response.Redirect(Application("baseroot")&"/editor/categorie/ListaCategorie.asp")	
		else
			Set catTmp = nothing
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=027")	
		end if
	else

		'**** verifico che non esista già una categoria con la stessa gerarchia
		catConflicts = false
		catIdTmp = objCategoria.findExsitingGerarchia(iGerarchia)
		if not(isNull(catTmp)) then
			catConflicts = true
		end if
		
		'**** verifico che non esista già una categoria con la stessa descrizione
		On Error Resume Next
		Set catIdTmp = objCategoria.findCategoriaByDescription(strDescrizione)
		if (strComp(typename(catIdTmp), "CategoryClass", 1) = 0) then
			catConflicts = true
		end if
		Set catIdTmp = nothing
		if(Err.number <>0)then
			catConflicts = false
		end if
	
		if not(catConflicts = true) then 		
			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans
			
			Dim idCatMax
			idCatMax = objCategoria.insertCategoria(iNumMenu, iGerarchia, strDescrizione, catType, bolContNews, bolContProd, bolVisible, idTemplate, metaDescription, metaKeyword, pageTitle, sub_domain_url, objConn)
			call objLogger.write("inserita categoria --> id: "&idCatMax&"; gerarchia: "&iGerarchia&"; descrizione: "&strDescrizione, objUserLogged.getUserName(), "info")
			
			if not(isNull(arrTarget)) then
				for each x in arrTarget
					call objCategoria.insertTargetXCategoria(x, idCatMax, objConn)
				next
			else
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=005")
			end if

			if (Cint(idTemplate) <> -1) then
				'*** RECUPERO LA LISTA DI LANGUAGE DISPONIBILI
				Set objLangTmp = New LanguageClass
				Set objLangList = objLangTmp.getListaLanguage()
				Set objLangTmp = nothing
					
				For Each x In objLangList
					lang_code_cat = UCase(objLangList(x).getLanguageDescrizione())
					id_lang_template = request("id_template_"&lang_code_cat)
					if (Cint(id_lang_template) <> -1) then
						call objCategoria.insertLangTemplateXCategoria(idCatMax, id_lang_template, lang_code_cat, objConn)
					end if
				Next		
				Set objLangList = nothing
			end if							
							
			if(catType = Application("strContentCat")) then
				targetType = 1
			elseif(catType = Application("strProdCat")) then
				targetType = 2
			else
				targetType = 4
			end if
			
			if(insert_new_target = 1) then
				if not(isNull(objTarget.findTargetByDescEq(old_target, objConn))) then						
					'select case targetType		
					'case 1	
						'if (strComp(typename(objTarget.findTargetByDescEqAndType(old_target, 1, objConn)), "TargetClass", 1) = 0) then
							'Set objTmpTarget = objTarget.findTargetByDescEqAndType(old_target, 1, objConn)
							'call objTarget.modifyTarget(objTmpTarget.getTargetID(), new_target, targetType, 0, 0, objConn)
							'Set objTmpTarget = nothing		
						'else
							'call objTarget.insertTarget(new_target, targetType, 0, 0, objConn)
						'end if
					'case 2	
						'if (strComp(typename(objTarget.findTargetByDescEqAndType(old_target, 2, objConn)), "TargetClass", 1) = 0) then
							'Set objTmpTarget = objTarget.findTargetByDescEqAndType(old_target, 2, objConn)
							'call objTarget.modifyTarget(objTmpTarget.getTargetID(), new_target, targetType, 0, 0, objConn)
							'Set objTmpTarget = nothing		
						'else
							'call objTarget.insertTarget(new_target, targetType, 0, 0, objConn)								
						'end if			
					'case 4
						'if (strComp(typename(objTarget.findTargetByDescEqAndType(old_target, 1, objConn)), "TargetClass", 1) = 0) then
							'Set objTmpTarget = objTarget.findTargetByDescEqAndType(old_target, 1, objConn)
							'call objTarget.modifyTarget(objTmpTarget.getTargetID(), new_target, 1, 0, 0, objConn)
							'Set objTmpTarget = nothing		
						'else
							'call objTarget.insertTarget(new_target, 1, 0, 0, objConn)							
						'end if
						'if (strComp(typename(objTarget.findTargetByDescEqAndType(old_target, 2, objConn)), "TargetClass", 1) = 0) then
							'Set objTmpTarget = objTarget.findTargetByDescEqAndType(old_target, 2, objConn)
							'call objTarget.modifyTarget(objTmpTarget.getTargetID(), new_target, 2, 0, 0, objConn)
							'Set objTmpTarget = nothing		
						'else
							'call objTarget.insertTarget(new_target, 2, 0, 0, objConn)									
						'end if					
					'case else
					'end select
				else
					select case targetType		
					case 1	
						call objTarget.insertTarget(new_target, targetType, 0, 0, objConn)
					case 2	
						call objTarget.insertTarget(new_target, targetType, 0, 0, objConn)			
					case 4
						call objTarget.insertTarget(new_target, 1, 0, 0, objConn)
						call objTarget.insertTarget(new_target, 2, 0, 0, objConn)				
					case else
					end select
				end if
			end if
			
			Dim objUsertmp, objListUsrTmp, usrTmp, objTargetTmp
			if(insert_new_target = 1 AND set_target_to_users = "1") then
				Set objUsertmp = new UserClass
				Set objListUsrTmp = objUsertmp.findUtente(null, "1,2", 1, null, 0, null)
			
				'call objLogger.write("typename --> "&typename(objListUsrTmp), "system", "debug")

				if (strComp(typename(objTarget.findTargetByDescEq(new_target, objConn)), "TargetClass", 1) = 0) then
					select case targetType		
					case 1	
						if (strComp(typename(objTarget.findTargetByDescEqAndType(new_target, 1, objConn)), "TargetClass", 1) = 0) then
							Set objTmpTarget = objTarget.findTargetByDescEqAndType(new_target, 1, objConn)
							for each x in objListUsrTmp
								Set usrTmp = objListUsrTmp(x)
								if(usrTmp.getRuolo() = Application("admin_role") OR usrTmp.getRuolo() = Application("editor_role")) then
									call objUsertmp.insertTargetXUser(objTmpTarget.getTargetID(), usrTmp.getUserID(), objConn)	
								end if					
								Set usrTmp = nothing		
							next
							
							if(set_target_to_categoria = 1) then
								call objCategoria.insertTargetXCategoria(objTmpTarget.getTargetID(), idCatMax, objConn)
							end if	
							Set objTmpTarget = nothing
						end if
					case 2	
						if (strComp(typename(objTarget.findTargetByDescEqAndType(new_target, 2, objConn)), "TargetClass", 1) = 0) then
							Set objTmpTarget = objTarget.findTargetByDescEqAndType(new_target, 2, objConn)
							for each x in objListUsrTmp
								Set usrTmp = objListUsrTmp(x)
								if(usrTmp.getRuolo() = Application("admin_role") OR usrTmp.getRuolo() = Application("editor_role")) then
									call objUsertmp.insertTargetXUser(objTmpTarget.getTargetID(), usrTmp.getUserID(), objConn)	
								end if					
								Set usrTmp = nothing		
							next
							
							if(set_target_to_categoria = 1) then
								call objCategoria.insertTargetXCategoria(objTmpTarget.getTargetID(), idCatMax, objConn)
							end if				
							Set objTmpTarget = nothing				
						end if			
					case 4
						if (strComp(typename(objTarget.findTargetByDescEqAndType(new_target, 1, objConn)), "TargetClass", 1) = 0) then
							Set objTmpTarget = objTarget.findTargetByDescEqAndType(new_target, 1, objConn)
							for each x in objListUsrTmp
								Set usrTmp = objListUsrTmp(x)
								if(usrTmp.getRuolo() = Application("admin_role") OR usrTmp.getRuolo() = Application("editor_role")) then
									call objUsertmp.insertTargetXUser(objTmpTarget.getTargetID(), usrTmp.getUserID(), objConn)	
								end if					
								Set usrTmp = nothing		
							next
							
							if(set_target_to_categoria = 1) then
								call objCategoria.insertTargetXCategoria(objTmpTarget.getTargetID(), idCatMax, objConn)
							end if	
							Set objTmpTarget = nothing						
						end if
						if (strComp(typename(objTarget.findTargetByDescEqAndType(new_target, 2, objConn)), "TargetClass", 1) = 0) then
							Set objTmpTarget = objTarget.findTargetByDescEqAndType(new_target, 2, objConn)
							for each x in objListUsrTmp
								Set usrTmp = objListUsrTmp(x)
								if(usrTmp.getRuolo() = Application("admin_role") OR usrTmp.getRuolo() = Application("editor_role")) then
									call objUsertmp.insertTargetXUser(objTmpTarget.getTargetID(), usrTmp.getUserID(), objConn)	
								end if					
								Set usrTmp = nothing		
							next
							
							if(set_target_to_categoria = 1) then
								call objCategoria.insertTargetXCategoria(objTmpTarget.getTargetID(), idCatMax, objConn)
							end if		
							Set objTmpTarget = nothing									
						end if				
					case else
					end select
				end if
				
				Set objListUsrTmp = nothing
				Set objUsertmp = nothing
			end if
						
			if objConn.Errors.Count = 0 then
				objConn.CommitTrans
			else
				objConn.RollBackTrans
				response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if			
			Set objDB = nothing			
			Set objCategoria = nothing
		
			response.Redirect(Application("baseroot")&"/editor/categorie/ListaCategorie.asp")				
		else
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=027")	
		end if
	end if
	
	Set objTarget = nothing
	Set objCategoria = nothing
	Set objLogger = nothing
	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>