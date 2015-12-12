<%
'<!--nsys-editinc3-->
%>
<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/NewsClass.asp" -->
<!-- #include virtual="/common/include/Objects/NewsletterClass.asp" -->
<!-- #include virtual="/common/include/Objects/ProductFieldClass.asp" -->
<%
'<!---nsys-editinc3-->
%>
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->
<%
Public Function convertDate(dateToConvert)
	Dim DD, MM, YY, HH, MIN, SS
	
	convertDate = null
	
	DD = DatePart("d", dateToConvert)
	MM = DatePart("m", dateToConvert)
	YY = DatePart("yyyy", dateToConvert)
	HH = DatePart("h", dateToConvert)
	MIN = DatePart("n", dateToConvert)
	SS = DatePart("s", dateToConvert)
	
	convertDate = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS		
End Function

if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if

	Dim field_name, field_val, objtype, id_objref
	field_name = request("field_name")
	field_val = request("field_val")
	objtype = request("objtype")
	id_objref = request("id_objref")

	Dim objRef, objTmp, objDict
	Select Case objtype
		Case "newsletter"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New NewsletterClass
			Set objTmp = objRef.findNewsletterByID(id_objref)
			objDict.add "descrizione",  objTmp.getDescrizione()
			objDict.add "stato", objTmp.getStato()
			objDict.add "template", objTmp.getTemplate()
			objDict.add "voucher", objTmp.getVoucher()
			Set objTmp = nothing
		
			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			call objRef.modifyNewsletter(id_objref, objDict.item("descrizione"), objDict.item("stato"), objDict.item("template"), objDict.item("voucher"))
			Set objRef = nothing
			Set objDict = nothing
		Case "content"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New NewsClass
			Set objTmp = objRef.findNewsByID(id_objref)
			objDict.add "titolo",  objTmp.getTitolo()
			objDict.add "abstract1", objTmp.getAbstract1()
			objDict.add "abstract2", objTmp.getAbstract2()
			objDict.add "abstract3", objTmp.getAbstract3()
			objDict.add "testo", objTmp.getTesto()
			objDict.add "keyword", objTmp.getKeyword()
			objDict.add "news_data", objTmp.getDataInsNews()
			objDict.add "news_data_pub", objTmp.getDataPubNews()
			objDict.add "news_data_del", objTmp.getDataDelNews()
			objDict.add "stato_news", objTmp.getStato()
			objDict.add "meta_description", objTmp.getMetaDescription()
			objDict.add "meta_keyword", objTmp.getMetaKeyword()
			objDict.add "page_title", objTmp.getPageTitle()			
			Set objTmp = nothing

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			'patch per formato date
			news_data_pub = objDict.item("news_data_pub")
			news_data_pub = convertDate(news_data_pub)
			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
			'objDict.remove("news_data_pub")
			'objDict.add "news_data_pub", news_data_pub
			objDict.Item("news_data_pub") = news_data_pub

			call objRef.modifyNewsNoTransaction(id_objref, objDict.item("titolo"), objDict.item("abstract1"), objDict.item("abstract2"), objDict.item("abstract3"), objDict.item("testo"), objDict.item("keyword"), objDict.item("news_data"), objDict.item("news_data_pub"), objDict.item("news_data_del"), objDict.item("stato_news"), objDict.item("meta_description"), objDict.item("meta_keyword"), objDict.item("page_title"))
			Set objRef = nothing
			Set objDict = nothing

			'rimuovo l'oggetto dalla cache
			Set objCacheClass = new CacheClass
			call objCacheClass.remove("content-"&id_objref)
			call objCacheClass.remove("listcf-"&id_objref)
			call objCacheClass.removeByPrefix("findc", id_objref)
			Set objCacheClass = nothing
		Case "content_field"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New ContentFieldClass
			Set objTmp = objRef.findContentFieldById(id_objref)
			objDict.add "description",  objTmp.getDescription()
			objDict.add "id_group", objTmp.getIdGroup()
			objDict.add "order", objTmp.getOrder()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			max_lenght = objTmp.getMaxLenght()
			if(isNull(max_lenght) OR max_lenght = "") then max_lenght = null end if
			if(isNull(objDict.item("id_group")) OR objDict.item("id_group") = "") then
				' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
				'objDict.remove("id_group")
				'objDict.add "id_group", null  
				objDict.Item("id_group") = null
			end if

			call objRef.modifyContentFieldNoTransaction(id_objref, objDict.item("description"), objDict.item("id_group"), objDict.item("order"), objTmp.getTypeField(), objTmp.getTypeContent(), objTmp.getRequired(), objTmp.getEnabled(), objTmp.getEditable(), max_lenght)			
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
	
			'rimuovo l'oggetto dalla cache
			Set objCacheClass = new CacheClass
			call objCacheClass.removeByPrefix("listcf-", null)
			Set objCacheClass = nothing
		Case "user"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New UserClass
			Set objTmp = objRef.findUserByID(id_objref)
			objDict.add "ruolo_utente",  objTmp.getRuolo()
			objDict.add "user_active", objTmp.getUserActive()
			objDict.add "public_profile", objTmp.getPublic()
			objDict.add "user_group", objTmp.getGroup()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			privacy = 0
			if(objTmp.getPrivacy()) then privacy = 1 end if
			newsletter = 0
			if(objTmp.getNewsletter()) then newsletter = 1 end if

			call objRef.modifyUserNoTransaction(id_objref, objTmp.getUserName(), null, objTmp.getEmail(), objDict.item("ruolo_utente"), privacy, newsletter, objDict.item("user_active"), objTmp.getSconto(), objTmp.getAdminComments(), objTmp.getInsertDate(), Now(), objDict.item("public_profile"), objDict.item("user_group"), objTmp.getAutomaticUser())
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "user_field"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New UserFieldClass
			Set objTmp = objRef.findUserFieldById(id_objref)
			objDict.add "description",  objTmp.getDescription()
			objDict.add "id_group", objTmp.getIdGroup()
			objDict.add "order", objTmp.getOrder()
			objDict.add "required", objTmp.getRequired()
			objDict.add "enabled", objTmp.getEnabled()
			objDict.add "use_for", objTmp.getUseFor()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			max_lenght = objTmp.getMaxLenght()
			if(isNull(max_lenght) OR max_lenght = "") then max_lenght = null end if
			if(isNull(objDict.item("id_group")) OR objDict.item("id_group") = "") then
				' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
				'objDict.remove("id_group")
				'objDict.add "id_group", null  
				objDict.Item("id_group") = null
			end if

			call objRef.modifyUserFieldNoTransaction(id_objref, objDict.item("description"), objDict.item("id_group"), objDict.item("order"), objTmp.getTypeField(), objTmp.getTypeContent(), objTmp.getValues(), objDict.item("required"), objDict.item("enabled"), max_lenght, objDict.item("use_for"))			
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "target"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New TargetClass
			Set objTmp = objRef.findTargetByID(id_objref)
			objDict.add "descrizione",  objTmp.getTargetDescrizione()
			objDict.add "target_type", objTmp.getTargetType()
			objDict.add "automatic", objTmp.isAutomatic()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			locked = 0
			if(objTmp.isLocked()) then locked = 1 end if

			call objRef.modifyTargetNoTransaction(id_objref, objDict.item("descrizione"), objDict.item("target_type"), locked, objDict.item("automatic"))			
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "category"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New CategoryClass
			Set objTmp = objRef.findCategoriaByID(id_objref)
			objDict.add "num_menu",  objTmp.getNumMenu()
			objDict.add "gerarchia", objTmp.getCatGerarchia()
			objDict.add "contiene_news", objTmp.contieneNews()
			objDict.add "contiene_prod", objTmp.contieneProd()
			objDict.add "visibile", objTmp.isCatVisible()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			call objRef.modifyCategoriaNoTransaction(id_objref, objDict.item("num_menu"), objDict.item("gerarchia"), objTmp.getCatDescrizione(), objTmp.getCatType(), objDict.item("contiene_news"), objDict.item("contiene_prod"), objDict.item("visibile"), objTmp.getIDTemplate(), objTmp.getMetaDescription(), objTmp.getMetaKeyword(), objTmp.getPageTitle(), objTmp.getSubDomainURL())		
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "template"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New TemplateClass
			Set objTmp = objRef.findTemplateByID(id_objref)
			objDict.add "descrizione_template",  objTmp.getDescrizioneTemplate()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans
			call objRef.modifyTemplate(id_objref, objTmp.getDirTemplate(), objTmp.getTemplateCss(), objDict.item("descrizione_template"), objTmp.getBaseTemplate(), objTmp.getOrderBy(), objTmp.getElemXPage(), objConn)			
			if objConn.Errors.Count = 0 AND Err.Number = 0 then
				objConn.CommitTrans
			else		
				objConn.RollBackTrans	
				'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if
			Set objDB = nothing	

			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "country"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New CountryClass
			Set objTmp = objRef.findCountryByID(id_objref)
			objDict.add "country_code",  objTmp.getCountryCode()
			objDict.add "state_region_code",  objTmp.getStateRegionCode()
			objDict.add "country_description",  objTmp.getCountryDescription()
			objDict.add "state_region_description",  objTmp.getStateRegionDescription()
			objDict.add "active",  objTmp.isActive()
			objDict.add "use_for",  objTmp.getUseFor()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			call objRef.modifyCountry(id_objref,objDict.item("country_code"), objDict.item("country_description"), objDict.item("state_region_code"), objDict.item("state_region_description"), objDict.item("active"), objDict.item("use_for"))		
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
'<!--nsys-editinc4-->
		Case "payment"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New PaymentClass
			Set objTmp = objRef.findPaymentByID(id_objref)
			objDict.add "descrizione",  objTmp.getDescrizione()
			objDict.add "dati_pagamento",  objTmp.getDatiPagamento()
			objDict.add "active",  objTmp.getAttivo()
			objDict.add "payment_type",  objTmp.getPaymentType()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans
			call objRef.modifyPayment(id_objref, objTmp.getKeywordMultilingua(), objDict.item("descrizione"), objDict.item("dati_pagamento"), objTmp.getCommission(), objTmp.getCommissionType(), objTmp.getURL(), objTmp.getPaymentModuleID(), objDict.item("active"), objDict.item("payment_type"), objConn)		
			if objConn.Errors.Count = 0 AND Err.Number = 0 then
				objConn.CommitTrans
			else		
				objConn.RollBackTrans	
				'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			Set objDB = nothing

			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "currency"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New CurrencyClass
			Set objTmp = objRef.findCurrencyByID(id_objref)
			objDict.add "attivo",  objTmp.getActive()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			default = 0
			if(objTmp.getDefault()) then default = 1 end if

			Dim DD, MM, YY, HH, MIN, SS
			
			DD = DatePart("d", objTmp.getDtaRefer())
			MM = DatePart("m", objTmp.getDtaRefer())
			YY = DatePart("yyyy", objTmp.getDtaRefer())	
			dtaReferer = YY&"-"&MM&"-"&DD
			
			DD = DatePart("d", objTmp.getDtaInsert())
			MM = DatePart("m", objTmp.getDtaInsert())
			YY = DatePart("yyyy", objTmp.getDtaInsert())
			HH = DatePart("h", objTmp.getDtaInsert())
			MIN = DatePart("n", objTmp.getDtaInsert())
			SS = DatePart("s", objTmp.getDtaInsert())	
			dtaInsert = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS

			rate = Replace(objTmp.getRate(), ".",",")

			call objRef.modifyCurrency(id_objref, objTmp.getCurrency(), rate, dtaReferer, dtaInsert, objDict.item("attivo"), default)		
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "tax"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New TaxsClass
			Set objTmp = objRef.findTassaByID(id_objref)
			objDict.add "descrizione",  objTmp.getDescrizioneTassa()
			objDict.add "valore",  objTmp.getValore()
			objDict.add "tipo_valore",  objTmp.getTipoValore()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			call objRef.modifyTassa(id_objref, objDict.item("descrizione"), objDict.item("valore"), objDict.item("tipo_valore"))		
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "taxs_group"		
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New TaxsGroupClass		
			composite_key = Split(id_objref, "|", -1, 1)
			id_group = composite_key(0)
			country_code = composite_key(1)
			state_region_code = composite_key(2)
			Set objTmp = objRef.findTaxsGroupValue(id_group, country_code, state_region_code)
			objDict.add "id_tassa_applicata",  objTmp.getTaxID()
			objDict.add "exclude_calculation",  objTmp.isExcludeCalculation()	

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"			
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val
			
			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans
			call objRef.modifyTaxsGroupValue(id_group,country_code, state_region_code, objDict.item("id_tassa_applicata"), objDict.item("exclude_calculation"), objConn)
			if objConn.Errors.Count = 0 AND Err.Number = 0 then
				objConn.CommitTrans
			else		
				objConn.RollBackTrans	
				'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			Set objDB = nothing
			
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "bill"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New BillsClass
			Set objTmp = objRef.findSpesaByID(id_objref)
			objDict.add "descrizione",  objTmp.getDescrizioneSpesa()
			objDict.add "valore",  objTmp.getValore()
			objDict.add "tipo_valore",  objTmp.getTipoValore()
			objDict.add "id_tassa_applicata",  objTmp.getIDTassaApplicata()
			objDict.add "applica_frontend",  objTmp.getApplicaFrontend()
			objDict.add "applica_backend",  objTmp.getApplicaBackend()
			objDict.add "taxs_group",  objTmp.getTaxGroup()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans
			call objRef.modifySpesa(id_objref, objDict.item("descrizione"), objDict.item("valore"), objDict.item("tipo_valore"), objDict.item("id_tassa_applicata"), objDict.item("applica_frontend"), objDict.item("applica_backend"), objTmp.getAutoactive(), objTmp.getMultiply(), objTmp.getRequired(), objTmp.getGroup(), objDict.item("taxs_group"), objTmp.getTypeView(), objConn)
			if objConn.Errors.Count = 0 AND Err.Number = 0 then
				objConn.CommitTrans
			else		
				objConn.RollBackTrans	
				'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			Set objDB = nothing

			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "margin"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New MarginDiscountClass
			Set objTmp = objRef.findMarginDiscountByID(id_objref)
			objDict.add "margine",  objTmp.getMargin()
			objDict.add "discount",  objTmp.getDiscount()
			objDict.add "prod_disc",  objTmp.isApplyProdDiscount()
			objDict.add "user_disc",  objTmp.isApplyUserDiscount()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans
			call objRef.modifyMarginDiscount(id_objref, objDict.item("margine"), objDict.item("discount"), objDict.item("prod_disc"), objDict.item("user_disc"), objConn)
			if objConn.Errors.Count = 0 AND Err.Number = 0 then
				objConn.CommitTrans
			else		
				objConn.RollBackTrans	
				'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			Set objDB = nothing
			
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "margin_group"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New UserGroupClass
			Set objTmp = objRef.findUserGroupByID(id_objref)
			objDict.add "short_desc",  objTmp.getShortDesc()
			objDict.add "long_desc",  objTmp.getLongDesc()
			objDict.add "taxs_group",  objTmp.getTaxGroup()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			call objRef.modifyUserGroup(id_objref, objDict.item("short_desc"), objDict.item("long_desc"), objTmp.isDefault(), objDict.item("taxs_group"))
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "business_rule"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New BusinessRulesClass
			Set objTmp = objRef.findRuleByID(id_objref)
			objDict.add "label",  objTmp.getLabel()
			objDict.add "descrizione",  objTmp.getDescrizione()
			objDict.add "activate",  objTmp.getActivate()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans
			
			
			call objRef.modifyRule(id_objref, objTmp.getRuleType(), objDict.item("label"), objDict.item("descrizione"), objDict.item("activate"), objTmp.getVoucherID(), objConn)
			if objConn.Errors.Count = 0 AND Err.Number = 0 then
				objConn.CommitTrans
			else		
				objConn.RollBackTrans	
				'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			Set objDB = nothing
			
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "product"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New ProductsClass
			Set objTmp = objRef.findProdottoByID(id_objref, 0)
			objDict.add "nome_prod",  objTmp.getNomeProdotto()
			objDict.add "stato_prod",  objTmp.getAttivo()
			objDict.add "prezzo_prod",  objTmp.getPrezzo()
			objDict.add "id_tassa_applicata",  objTmp.getIDTassaApplicata()
			objDict.add "meta_description", objTmp.getMetaDescription()
			objDict.add "meta_keyword", objTmp.getMetaKeyword()
			objDict.add "page_title", objTmp.getPageTitle()
			objDict.add "edit_buy_qta", objTmp.getEditBuyQta()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			field_prezzo_prod = objDict.item("prezzo_prod")
			field_prezzo_prod = Replace(field_prezzo_prod, ".", "", 1, -1, 1)
			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
			'objDict.remove("prezzo_prod")
			'objDict.add "prezzo_prod", field_prezzo_prod
			objDict.Item("prezzo_prod") = field_prezzo_prod

			call objRef.modifyProdottoNoTransaction(id_objref, objDict.item("nome_prod"), objTmp.getSommarioProdotto(), objTmp.getDescProdotto(), objDict.item("prezzo_prod"), objTmp.getQtaDisp(), objDict.item("stato_prod"), objTmp.getSconto(), objTmp.getCodiceProd(), objDict.item("id_tassa_applicata"), objTmp.getProdType(), objTmp.getMaxDownload(), objTmp.getMaxDownloadTime(), objTmp.getTaxGroup(), objDict.item("meta_description"), objDict.item("meta_keyword"), objDict.item("page_title"), objDict.item("edit_buy_qta"))
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
			
			'rimuovo l'oggetto dalla cache
			Set objCacheClass = new CacheClass
			Set objBase64 = new Base64Class
			objCacheClass.remove("product-"&objBase64.Base64Encode(id_objref))
			call objCacheClass.removeByPrefix("findp", id_objref)
			Set objBase64 = nothing
			Set objCacheClass = nothing
		Case "product_field"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New ProductFieldClass
			Set objTmp = objRef.findProductFieldById(id_objref)
			objDict.add "description",  objTmp.getDescription()
			objDict.add "id_group", objTmp.getIdGroup()
			objDict.add "order", objTmp.getOrder()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			max_lenght = objTmp.getMaxLenght()
			if(isNull(max_lenght) OR max_lenght = "") then max_lenght = null end if
			if(isNull(objDict.item("id_group")) OR objDict.item("id_group") = "") then
				' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
				'objDict.remove("id_group")
				'objDict.add "id_group", null  
				objDict.Item("id_group") = null
			end if

			call objRef.modifyProductFieldNoTransaction(id_objref, objDict.item("description"), objDict.item("id_group"), objDict.item("order"), objTmp.getTypeField(), objTmp.getTypeContent(), objTmp.getRequired(), objTmp.getEnabled(), objTmp.getEditable(), max_lenght)			
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
		Case "voucher"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New VoucherClass
			Set objTmp = objRef.findCampaignByID(id_objref)
			objDict.add "label",  objTmp.getLabel()
			objDict.add "valore",  objTmp.getValore()
			objDict.add "operation",  objTmp.getOperation()
			objDict.add "activate",  objTmp.getActivate()
			objDict.add "exlude_prod_rule",  objTmp.getExcludeProdRule()

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			Set objDB = New DBManagerClass
			Set objConn = objDB.openConnection()
			objConn.BeginTrans
			
			call objRef.modifyCampaign(id_objref, objDict.item("label"), objTmp.getVoucherType(), objTmp.getDescrizione(), objDict.item("valore"), objDict.item("operation"), objDict.item("activate"), objTmp.getMaxGeneration(), objTmp.getMaxUsage(), objTmp.getEnableDate(), objTmp.getExpireDate(), objDict.item("exlude_prod_rule"), objConn)
			if objConn.Errors.Count = 0 AND Err.Number = 0 then
				objConn.CommitTrans
			else		
				objConn.RollBackTrans	
				'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			Set objDB = nothing
			
			Set objTmp = nothing
			Set objRef = nothing
			Set objDict = nothing
'<!---nsys-editinc4-->
		Case Else			
	End Select
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>
