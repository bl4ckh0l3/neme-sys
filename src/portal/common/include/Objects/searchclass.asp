<%

Class SearchClass

	private skipWord	

	Private Sub Class_Initialize()
		Set skipWord = Server.CreateObject("Scripting.Dictionary")
		skipWord.add "a","" 
		skipWord.add "all","" 
		skipWord.add "am","" 
		skipWord.add "an","" 
		skipWord.add "and","" 
		skipWord.add "any","" 
		skipWord.add "are","" 
		skipWord.add "as","" 
		skipWord.add "at","" 
		skipWord.add "be","" 
		skipWord.add "but","" 
		skipWord.add "can","" 
		skipWord.add "did","" 
		skipWord.add "do","" 
		skipWord.add "does","" 
		skipWord.add "for","" 
		skipWord.add "from","" 
		skipWord.add "had","" 
		skipWord.add "has","" 
		skipWord.add "have","" 
		skipWord.add "here","" 
		skipWord.add "how","" 
		skipWord.add "i","" 
		skipWord.add "if","" 
		skipWord.add "in","" 
		skipWord.add "is","" 
		skipWord.add "it","" 
		skipWord.add "no","" 
		skipWord.add "not","" 
		skipWord.add "of","" 
		skipWord.add "on","" 
		skipWord.add "or","" 
		skipWord.add "so","" 
		skipWord.add "that","" 
		skipWord.add "the","" 
		skipWord.add "then","" 
		skipWord.add "there","" 
		skipWord.add "this","" 
		skipWord.add "to","" 
		skipWord.add "too","" 
		skipWord.add "up","" 
		skipWord.add "use","" 
		skipWord.add "what","" 
		skipWord.add "when","" 
		skipWord.add "where","" 
		skipWord.add "who","" 
		skipWord.add "why","" 
		skipWord.add "you",""
		skipWord.add "di",""
		skipWord.add "del",""
		skipWord.add "dell'",""
		skipWord.add "dello",""
		skipWord.add "della",""
		skipWord.add "dei",""
		skipWord.add "degli",""
		skipWord.add "delle",""
		skipWord.add "al",""
		skipWord.add "all'",""
		skipWord.add "allo",""
		skipWord.add "alla",""
		skipWord.add "ai",""
		skipWord.add "agli",""
		skipWord.add "alle",""
		skipWord.add "da",""
		skipWord.add "dal",""
		skipWord.add "dall'",""
		skipWord.add "dallo",""
		skipWord.add "dalla",""
		skipWord.add "dai",""
		skipWord.add "dagli",""
		skipWord.add "dalle",""
		skipWord.add "nel",""
		skipWord.add "nell'",""
		skipWord.add "nello",""
		skipWord.add "nella",""
		skipWord.add "nei",""
		skipWord.add "negli",""
		skipWord.add "nelle",""
		skipWord.add "su",""
		skipWord.add "sul",""
		skipWord.add "sull'",""
		skipWord.add "sullo",""
		skipWord.add "sulla",""
		skipWord.add "sui",""
		skipWord.add "sugli",""
		skipWord.add "sulle",""
		skipWord.add "con",""
		skipWord.add "col",""
		skipWord.add "coll'",""
		skipWord.add "collo",""
		skipWord.add "colla",""
		skipWord.add "coi",""
		skipWord.add "cogli",""
		skipWord.add "colle",""
		skipWord.add "per",""
		skipWord.add "pel",""
		skipWord.add "pei",""
		skipWord.add "fra",""
		skipWord.add "tra",""
		skipWord.add "il",""
		skipWord.add "lo",""
		skipWord.add "l'",""
		skipWord.add "el",""
		skipWord.add "o",""
		skipWord.add "u",""
		skipWord.add "le",""
		skipWord.add "e",""
		skipWord.add "gli",""
		skipWord.add "los",""
		skipWord.add "els",""
		skipWord.add "os",""
		skipWord.add "les",""
		skipWord.add "la",""
		skipWord.add "las",""
		skipWord.add "un",""
		skipWord.add "uno",""
		skipWord.add "um",""
		skipWord.add "unui",""
		skipWord.add "unos",""
		skipWord.add "uns",""
		skipWord.add "unor",""
		skipWord.add "una",""
		skipWord.add "un'",""
		skipWord.add "uma",""
		skipWord.add "une",""
		skipWord.add "unei",""
		skipWord.add "unas",""
		skipWord.add "unes",""
		skipWord.add "umas",""
	End Sub
	
	Private Sub Class_Terminate()
		Set skipWord = nothing
	End Sub

	'****************** funzione di recupero delle news tramite vari parametri		
	Public Function searchNews(id, titolo, abstract1, abstract2, abstract3, text, keyword, arrTargetCat, arrTargetLang, data_pub, data_del, stato_news, order_by, do_And, bolAddFiles)
		searchNews = null				
		Dim objDB, strSQL, strSQLTarget, strSQLTmp, objRS, objRSTarget, objConn, srtAndOr
		srtAndOr = " OR "
		if(Cbool(do_And) = true) then
			srtAndOr = " AND "
		end if		

		Dim hasTarget
		hasTarget = true
		Dim noTargetCat,noTargetlang
		noTargetCat = false
		noTargetlang = false
		
		noTargetCat = (isNull(arrTargetCat) OR not(strComp(typename(arrTargetCat), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0))
		
		if (noTargetCat) AND (noTargetlang) then
			hasTarget = false
		end if	

		strSQL = "SELECT * FROM news_find"
		if (isNull(id) AND isNull(titolo) AND isNull(abstract1) AND isNull(abstract2) AND isNull(abstract3) AND isNull(text) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news) AND not(hasTarget)) then
			strSQL = "SELECT * FROM news_find"
		else
			strSQL = strSQL & " WHERE"	

			if not(isNull(id)) OR not(isNull(titolo)) OR not(isNull(abstract1)) OR not(isNull(abstract2)) OR not(isNull(abstract3)) OR not(isNull(text)) OR not(isNull(keyword)) then
				strSQL = strSQL & "("
				if not(isNull(id)) then strSQL = strSQL & srtAndOr & "id=?"
				
				Dim arrSeparator
				arrSeparator = Array("  ", " ", "; ", ", ", ": ")
				
				if not(isNull(titolo)) then
					for each xSep in arrSeparator
						on Error Resume Next
						titolo = Split(Trim(titolo), xSep, -1, 1)
						titolo = join(titolo)
						if(Err.number<>0) then
						end if
					next
					titolo = Split(Trim(titolo), " ", -1, 1)
					
					for each x in titolo
						if not(skipWord.Exists(x)) then
						strSQL = strSQL & srtAndOr & "titolo LIKE ?"
						end if
					next
				end if
				if not(isNull(abstract1)) then	
					for each xSep in arrSeparator	
						on Error Resume Next
						abstract1 = Split(Trim(abstract1), xSep, -1, 1)
						abstract1 = join(abstract1)
						if(Err.number<>0) then
						end if
					next	
					abstract1 = Split(Trim(abstract1), " ", -1, 1)

					for each x in abstract1
						if not(skipWord.Exists(x)) then 
						strSQL = strSQL & srtAndOr & "abstract LIKE ?"
						end if
					next
				end if	
				if not(isNull(abstract2)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						abstract2 = Split(Trim(abstract2), xSep, -1, 1)
						abstract2 = join(abstract2)
						if(Err.number<>0) then
						end if
					next
					abstract2 = Split(Trim(abstract2), " ", -1, 1)
					
					for each x in abstract2 
						if not(skipWord.Exists(x)) then
						strSQL = strSQL & srtAndOr & "abstract_2 LIKE ?"
						end if
					next
				end if				
				if not(isNull(abstract3)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						abstract3 = Split(Trim(abstract3), xSep, -1, 1)
						abstract3 = join(abstract3)
						if(Err.number<>0) then
						end if
					next
					abstract3 = Split(Trim(abstract3), " ", -1, 1)
					
					for each x in abstract3 
						if not(skipWord.Exists(x)) then
						strSQL = strSQL & srtAndOr & "abstract_3 LIKE ?"
						end if
					next
				end if				
				if not(isNull(text)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						text = Split(Trim(text), xSep, -1, 1)
						text = join(text)
						if(Err.number<>0) then
						end if
					next
					text = Split(Trim(text), " ", -1, 1)
					
					for each x in text
						if not(skipWord.Exists(x)) then 
						strSQL = strSQL & srtAndOr & "testo LIKE ?"
						end if
					next
				end if				
				if not(isNull(keyword)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						keyword = Split(Trim(keyword), xSep, -1, 1)
						keyword = join(keyword)
						if(Err.number<>0) then
						end if
					next
					keyword = Split(Trim(keyword), " ", -1, 1)
					
					for each x in keyword
						if not(skipWord.Exists(x)) then 
						strSQL = strSQL & srtAndOr & "keyword LIKE ?"
						end if
					next
				end if				
				
				if not(isNull(data_pub)) then 
					if (Application("dbType") = 1) then
						strSQL = strSQL & srtAndOr & "data_pubblicazione <=?" 	
					else
						strSQL = strSQL & srtAndOr & "data_pubblicazione <=#?#" 						
					end if
				end if
				if not(isNull(data_del)) then
					if (Application("dbType") = 1) then
						strSQL = strSQL & srtAndOr & "(data_cancellazione >=? OR data_cancellazione='0000-00-00 00:00:00')" 	
					else
						strSQL = strSQL & srtAndOr & "data_cancellazione >=#?#" 						
					end if
				end if
				strSQL = strSQL & ")"
			end if
			if not(isNull(stato_news)) then strSQL = strSQL & " AND stato_news=?"			
			if (hasTarget) then 
				strSQL = strSQL & " AND id IN("				
				strSQL = strSQL & "SELECT DISTINCT(id_news) FROM target_x_news WHERE"
				if not(noTargetCat) then
					if(arrTargetCat.Count > 0) then
						strSQL = strSQL & " id_news IN(SELECT DISTINCT(id_news) FROM target_x_news WHERE id_target IN("								
						for each idx in arrTargetCat
							strSQL = strSQL &idx&","
						next					
						strSQL = strSQL & "))"
					end if
				end if
				if not(noTargetLang) then
					if(arrTargetLang.Count > 0) then
						strSQL = strSQL & " AND id_target IN("	
						for each idy in arrTargetLang
							strSQL = strSQL &idy&","
						next					
						strSQL = strSQL & ")"	
					end if
				end if
				strSQL = strSQL & ")"
				strSQL = Replace(strSQL, ",)", ")", 1, -1, 1)
				strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
				strSQL = Trim(strSQL)
			end if
		end if

		if not(isNull(order_by)) then
			select Case order_by
			Case 1
				strSQL = strSQL & " ORDER BY titolo ASC"
			Case 2
				strSQL = strSQL & " ORDER BY titolo DESC"
			Case 3
				strSQL = strSQL & " ORDER BY abstract ASC"
			Case 4
				strSQL = strSQL & " ORDER BY abstract DESC"
			Case 5
				strSQL = strSQL & " ORDER BY abstract_2 ASC"
			Case 6
				strSQL = strSQL & " ORDER BY abstract_2 DESC"
			Case 7
				strSQL = strSQL & " ORDER BY abstract_3 ASC"
			Case 8
				strSQL = strSQL & " ORDER BY abstract_3 DESC"
			Case 9
				strSQL = strSQL & " ORDER BY testo ASC"
			Case 10
				strSQL = strSQL & " ORDER BY testo DESC"
			Case 11
				strSQL = strSQL & " ORDER BY data_pubblicazione ASC"
			Case 12
				strSQL = strSQL & " ORDER BY data_pubblicazione DESC"
			Case 13
				strSQL = strSQL & " ORDER BY data_inserimento ASC"
			Case 14
				strSQL = strSQL & " ORDER BY data_inserimento DESC"
			Case Else
			End Select
		else
			strSQL = strSQL & " ORDER BY titolo ASC"
		end if

		strSQL = Replace(strSQL, "WHERE( AND", "WHERE(", 1, -1, 1)
		strSQL = Replace(strSQL, "WHERE( OR", "WHERE(", 1, -1, 1)
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Replace(strSQL, "WHERE OR", "WHERE", 1, -1, 1)
		strSQL = strSQL & ";"
		strSQL = Trim(strSQL)	

		Dim objCommand		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		if (isNull(id) AND isNull(titolo) AND isNull(abstract1) AND isNull(abstract2) AND isNull(abstract3) AND isNull(text) AND isNull(keyword) AND isNull(data_pub) AND isNull(data_del) AND isNull(stato_news)) then
		else
			if not(isNull(id)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)

			if not(isNull(titolo)) AND IsArray(titolo)  then					
				for each x in titolo
					if not(skipWord.Exists(x)) then  
					objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,"%"&x&"%")
					end if
				next
			end if
			if not(isNull(abstract1)) AND IsArray(abstract1) then										
				for each x in abstract1
					if not(skipWord.Exists(x)) then  
					objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&x&"%")
					end if
				next
			end if				
			if not(isNull(abstract2)) AND IsArray(abstract2) then					
				for each x in abstract2
					if not(skipWord.Exists(x)) then 
					objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&x&"%")
					end if
				next
			end if				
			if not(isNull(abstract3)) AND IsArray(abstract3) then					
				for each x in abstract3 
					if not(skipWord.Exists(x)) then 
					objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&x&"%")
					end if
				next
			end if				
			if not(isNull(text)) AND IsArray(text) then					
				for each x in text
					if not(skipWord.Exists(x)) then  
					objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&x&"%")
					end if
				next
			end if				
			if not(isNull(keyword)) AND IsArray(keyword) then					
				for each x in keyword
					if not(skipWord.Exists(x)) then  
					objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,"%"&x&"%")
					end if
				next
			end if
			
			'il passaggio seguente � da verificare con query secca di test su DB
			if not(isNull(data_pub)) then 
				if (Application("dbType") = 1) then
					data_pub = convertDate(data_pub)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 	
				else
					data_pub = FormatDateTime(data_pub, 2)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_pub) 						
				end if
			end if
			if not(isNull(data_del)) then
				if (Application("dbType") = 1) then
					data_del = convertDate(data_del)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del) 	
				else
					data_pub = FormatDateTime(data_del, 2)
					objCommand.Parameters.Append objCommand.CreateParameter(,135,1,,data_del)  						
				end if
			end if
			if not(isNull(stato_news)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,stato_news)
		end if
		Set objRS = objCommand.Execute()
		
		if objRS.EOF then				
			searchNews = null			
		else			
			Dim objNews, objListaNews			
			Set objListaNews = Server.CreateObject("Scripting.Dictionary")

			do while not objRS.EOF			
				Dim id_news	
				id_news = objRS("id")
						
				Set objNews = new NewsClass				
				objNews.setNewsID(id_news)
				objNews.setTitolo(objRS("titolo"))
				objNews.setAbstract1(objRS("abstract"))
				objNews.setAbstract2(objRS("abstract_2"))
				objNews.setAbstract3(objRS("abstract_3"))
				objNews.setTesto(objRS("testo"))
				objNews.setDataInsNews(objRS("data_inserimento"))
				objNews.setDataPubNews(objRS("data_pubblicazione"))
				objNews.setDataDelNews(objRS("data_cancellazione"))
				objNews.setStato(objRS("stato_news"))
				objNews.setEditorID(objRS("id_utente"))
				objListaNews.add id_news, objNews
				Set objNews = Nothing
				objRS.moveNext()				
			loop
				
			if (objListaNews.Count > 0) then
				Dim objListaTarget, objFiles, objListaFiles		
				Set objNews = new NewsClass			
				Set objFiles = new File4NewsClass
				Set searchNews = Server.CreateObject("Scripting.Dictionary")
				Dim bolAddNews
				
				for each xNews in objListaNews
					on error resume next
					bolAddNews = true					
					Set objTargetCatPageTempl = getTargetCatPageTempl(xNews, 1)
					if not(Instr(1, typename(objTargetCatPageTempl), "Dictionary", 1) > 0) then
						bolAddNews = false
					end if					
						
					if(bolAddFiles)then
						Set objListaFiles = objFiles.getFilePerNews(xNews)				
						if (Instr(1, typename(objListaFiles), "Dictionary", 1) > 0) then
							objListaNews(xNews).setFilePerNews(objListaFiles)
						end if
						Set objListaFiles = nothing
					end if
		
					if Err.number <> 0 then
						bolAddNews = false
					end if							
				
					'imposto come chiave dell'oggetto Dictionary l'id della news  unita alla gerarchia della categoria trovata in base al target scelto della news e al page_num del template associato!
					'durante la fase di recupero bisogna splittare la chiave per recuperare solo la gerarchia e page_num
					'(questo espediente serve quando ci sono piu news che hanno la stessa gerarchia come chiave altrimenti alcuni risultati vengono eliminati dalla lista)
					if(bolAddNews) then
						searchNews.add xNews & "|" & objTargetCatPageTempl("gerarchia")&"-"&objTargetCatPageTempl("page_num"), objListaNews(xNews)
					end if
					Set objTargetCatPageTempl = nothing					
				next
				
				Set objFiles = nothing
				Set objNews = nothing
			else
				searchNews = null				
			end if
			
			Set objListaNews = nothing
		end if	
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing		
	End Function	
	
	'****************** funzione di recupero dei prodotti tramite vari parametri
	Public Function searchProduct(id, nome_prod, sommario_prod, desc_prod, prezzo, qta_disp, arrTargetCat, arrTargetLang, codice_prod, attivo, order_by, do_And, bolAddFiles)		
		searchProduct = null				
		Dim objDB, strSQL, strSQLTarget, strSQLTmp, objRS, objRSTarget, objConn, srtAndOr
		srtAndOr = " OR "
		if(Cbool(do_And) = true) then
			srtAndOr = " AND "
		end if		
		
		Dim hasTarget
		hasTarget = true
		Dim noTargetCat,noTargetlang
		noTargetCat = false
		noTargetlang = false
		
		noTargetCat = (isNull(arrTargetCat) OR not(strComp(typename(arrTargetCat), "Dictionary", 1) = 0))
		noTargetlang = (isNull(arrTargetLang) OR not(strComp(typename(arrTargetLang), "Dictionary", 1) = 0))
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		Set objListTarget = Server.CreateObject("Scripting.Dictionary")
		if (noTargetCat) AND (noTargetlang) then
			hasTarget = false
		end if
		
		strSQL = "SELECT * FROM prodotti"
		if (isNull(id) AND isNull(nome_prod) AND isNull(sommario_prod) AND isNull(desc_prod) AND isNull(prezzo) AND isNull(qta_disp) AND isNull(codice_prod) AND isNull(attivo) AND not(hasTarget)) then
			strSQL = "SELECT * FROM prodotti"
		else
			strSQL = strSQL & " WHERE"			
			if not(isNull(id)) OR not(isNull(nome_prod)) OR not(isNull(sommario_prod)) OR not(isNull(desc_prod)) OR not(isNull(prezzo)) OR not(isNull(qta_disp)) then
				strSQL = strSQL & "("
				if not(isNull(id)) then strSQL = strSQL & srtAndOr & "id_prodotto=?"
				
				Dim arrSeparator
				arrSeparator = Array("  ", " ", "; ", ", ", ": ")
				
				if not(isNull(nome_prod)) then
					for each xSep in arrSeparator
						on Error Resume Next
						nome_prod = Split(Trim(nome_prod), xSep, -1, 1)
						nome_prod = join(nome_prod)
						if(Err.number<>0) then
						end if
					next
					nome_prod = Split(Trim(nome_prod), " ", -1, 1)
					
					for each x in nome_prod
						if not(skipWord.Exists(x)) then  
						strSQL = strSQL & srtAndOr & "nome_prod LIKE ?"
						end if
					next
				end if
				if not(isNull(sommario_prod)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						sommario_prod = Split(Trim(sommario_prod), xSep, -1, 1)
						sommario_prod = join(sommario_prod)
						if(Err.number<>0) then
						end if
					next
					sommario_prod = Split(Trim(sommario_prod), " ", -1, 1)
					
					for each x in sommario_prod 
						if not(skipWord.Exists(x)) then 
						strSQL = strSQL & srtAndOr & "sommario_prod LIKE ?"
						end if
					next
				end if				
				if not(isNull(desc_prod)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						desc_prod = Split(Trim(desc_prod), xSep, -1, 1)
						desc_prod = join(desc_prod)
						if(Err.number<>0) then
						end if
					next
					desc_prod = Split(Trim(desc_prod), " ", -1, 1)
					
					for each x in desc_prod 
						if not(skipWord.Exists(x)) then 
						strSQL = strSQL & srtAndOr & "desc_prod LIKE ?"
						end if
					next
				end if				
				if not(isNull(prezzo)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						prezzo = Split(Trim(prezzo), xSep, -1, 1)
						prezzo = join(prezzo)
						if(Err.number<>0) then
						end if
					next
					prezzo = Split(Trim(prezzo), " ", -1, 1)
					
					for each x in prezzo
						if not(skipWord.Exists(x)) then  
						strSQL = strSQL & srtAndOr & "prezzo =?"
						end if
					next
				end if				
				if not(isNull(qta_disp)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						qta_disp = Split(Trim(qta_disp), xSep, -1, 1)
						qta_disp = join(qta_disp)
						if(Err.number<>0) then
						end if
					next
					qta_disp = Split(Trim(qta_disp), " ", -1, 1)
					
					for each x in qta_disp 
						if not(skipWord.Exists(x)) then 
						strSQL = strSQL & srtAndOr & "qta_disp =?"
						end if
					next
				end if					
				if not(isNull(codice_prod)) then					
					for each xSep in arrSeparator
						on Error Resume Next
						codice_prod = Split(Trim(codice_prod), xSep, -1, 1)
						codice_prod = join(codice_prod)
						if(Err.number<>0) then
						end if
					next
					codice_prod = Split(Trim(codice_prod), " ", -1, 1)
					
					for each x in codice_prod
						if not(skipWord.Exists(x)) then  
						strSQL = strSQL & srtAndOr & "codice_prod =?"
						end if
					next
				end if				
				
				strSQL = strSQL & ")"
			end if
			if not(isNull(attivo)) then strSQL = strSQL & " AND " & "attivo=?"
			if (hasTarget) then 
				strSQL = strSQL & " AND id_prodotto IN("				
				strSQL = strSQL & "SELECT DISTINCT(id_prodotto) FROM target_x_prodotto WHERE"
				if not(noTargetCat) then
					if(arrTargetCat.Count > 0) then
						strSQL = strSQL & " id_news IN(SELECT DISTINCT(id_prodotto) FROM target_x_prodotto WHERE id_target IN("								
						for each idx in arrTargetCat
							strSQL = strSQL &idx&","
						next					
						strSQL = strSQL & "))"
					end if
				end if
				if not(noTargetLang) then
					if(arrTargetLang.Count > 0) then
						strSQL = strSQL & " AND id_target IN("	
						for each idy in arrTargetLang
							strSQL = strSQL &idy&","
						next					
						strSQL = strSQL & ")"	
					end if
				end if
				strSQL = strSQL & ")"
				strSQL = Replace(strSQL, ",)", ")", 1, -1, 1)
				strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
				strSQL = Trim(strSQL)
			end if
		end if
		
		if not(isNull(order_by)) then
			select Case order_by
			Case 1
				strSQL = strSQL & " ORDER BY nome_prod ASC"
			Case 2
				strSQL = strSQL & " ORDER BY nome_prod DESC"
			Case 3
				strSQL = strSQL & " ORDER BY sommario_prod ASC"
			Case 4
				strSQL = strSQL & " ORDER BY sommario_prod DESC"
			Case 5
				strSQL = strSQL & " ORDER BY desc_prod ASC"
			Case 6
				strSQL = strSQL & " ORDER BY desc_prod DESC"
			Case 7
				strSQL = strSQL & " ORDER BY prezzo ASC"
			Case 8
				strSQL = strSQL & " ORDER BY prezzo DESC"
			Case 9
				strSQL = strSQL & " ORDER BY qta_disp ASC"
			Case 10
				strSQL = strSQL & " ORDER BY qta_disp DESC"
			Case 11
				strSQL = strSQL & " ORDER BY codice_prod ASC"
			Case 12
				strSQL = strSQL & " ORDER BY codice_prod DESC"
			Case Else
			End Select
		else
			strSQL = strSQL & " ORDER BY prezzo ASC"
		end if

		strSQL = Replace(strSQL, "WHERE( AND", "WHERE(", 1, -1, 1)
		strSQL = Replace(strSQL, "WHERE( OR", "WHERE(", 1, -1, 1)
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Replace(strSQL, "WHERE OR", "WHERE", 1, -1, 1)
		strSQL = strSQL & ";"
		strSQL = Trim(strSQL)

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		
		if (isNull(id) AND isNull(nome_prod) AND isNull(sommario_prod) AND isNull(desc_prod) AND isNull(prezzo) AND isNull(qta_disp) AND isNull(codice_prod) AND isNull(attivo)) then
		else
			if not(isNull(id)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
			if not(isNull(nome_prod)) AND IsArray(nome_prod)  then					
				for each x in nome_prod 
					if not(skipWord.Exists(x)) then
					objCommand.Parameters.Append objCommand.CreateParameter(,200,1,250,"%"&x&"%")
					end if
				next
			end if
			if not(isNull(sommario_prod)) AND IsArray(sommario_prod) then										
				for each x in sommario_prod
					if not(skipWord.Exists(x)) then
					objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&x&"%")
					end if
				next
			end if				
			if not(isNull(desc_prod)) AND IsArray(desc_prod) then					
				for each x in desc_prod 
					if not(skipWord.Exists(x)) then
					objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,"%"&x&"%")
					end if
				next
			end if				
			if not(isNull(prezzo)) AND IsArray(prezzo) then					
				for each x in prezzo 
					if not(skipWord.Exists(x)) then
					objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(x)))
					end if
				next
			end if				
			if not(isNull(qta_disp)) AND IsArray(qta_disp) then					
				for each x in qta_disp 
					if not(skipWord.Exists(x)) then
					objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,Trim(x))
					end if
				next
			end if				
			if not(isNull(codice_prod)) AND IsArray(codice_prod) then					
				for each x in codice_prod
					if not(skipWord.Exists(x)) then 
					objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,Trim(x))
					end if
				next
			end if
			if not(isNull(attivo)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,attivo)
		end if
		Set objRS = objCommand.Execute()
				
		if objRS.EOF then				
			searchProduct = null			
		else			
			Dim objProd, objListaProd			
			Set objListaProd = Server.CreateObject("Scripting.Dictionary")			
			
			do while not objRS.EOF			
				Dim id_prod	
				id_prod = objRS("id_prodotto")
							
				Set objProd = new ProductsClass				
				objProd.setIDProdotto(objRS("id_prodotto"))    
				objProd.setNomeProdotto(objRS("nome_prod"))
				objProd.setSommarioProdotto(objRS("sommario_prod"))
				objProd.setDescProdotto(objRS("desc_prod"))
				objProd.setPrezzo(objRS("prezzo"))
				objProd.setQtaDisp(objRS("qta_disp"))
				objProd.setAttivo(objRS("attivo"))
				objProd.setSconto(objRS("sconto"))
				objProd.setCodiceProd(objRS("codice_prod"))
				objListaProd.add id_prod, objProd
				Set objProd = Nothing
				objRS.moveNext()				
			loop
			
			if (objListaProd.Count > 0) then
				Dim objListaTarget, objFiles, objListaFiles		
				Set objProd = new ProductsClass			
				Set objFiles = new File4ProductsClass
				Set searchProduct = Server.CreateObject("Scripting.Dictionary")
				Dim bolAddProd
				
				for each xProd in objListaProd
					on error resume next
					bolAddProd = true	

					Set objTargetCatPageTempl = getTargetCatPageTempl(xProd, 2)
					if not(Instr(1, typename(objTargetCatPageTempl), "Dictionary", 1) > 0) then
						bolAddProd = false
					end if
						
					if(bolAddFiles)then
						Set objListaFiles = objFiles.getFilePerProdotto(xProd)				
						if (Instr(1, typename(objListaFiles), "Dictionary", 1) > 0) then
							objListaProd(xProd).setFileXProdotto(objListaFiles)
						end if
						Set objListaFiles = nothing
					end if					

					if Err.number <> 0 then
						bolAddProd = false
					end if							
				
					'imposto come chiave dell'oggetto Dictionary l'id del prod unita alla gerarchia della categoria trovata in base al target scelto del prod e al page_num del template associato!
					'durante la fase di recupero bisogna splittare la chiave per recuperare solo la gerarchia e page_num
					'(questo espediente serve quando ci sono piu prod che hanno la stessa gerarchia come chiave altrimenti alcuni risultati vengono eliminati dalla lista)
					if(bolAddProd) then
						searchProduct.add xProd & "|" & objTargetCatPageTempl("gerarchia")&"-"&objTargetCatPageTempl("page_num"), objListaProd(xProd)
					end if
					Set objTargetCatPageTempl = nothing
				next
				
				Set objFiles = nothing
				Set objProd = nothing
			else
				searchProduct = null				
			end if
			
			Set objListaProd = nothing
		end if		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing		
	End Function	
	
	'****************** funzione di recupero di news e prodotti tramite vari parametri
	Public Function searchAll(text, arrTargetLang, bolAddFiles)		
		Set searchAll = Server.CreateObject("Scripting.Dictionary")
		bolHasNews = false
		bolHasProd = false

		on error resume next
		Set objListaNews = searchNews(null, text, text, text, text, text, text, null, arrTargetLang, null, null, 1, 1, false, bolAddFiles)
		if (objListaNews.Count > 0) then
			bolHasNews = true
		end if
		if Err.number <> 0 then
			'response.write(Err.description)
			bolHasNews = false
		end if		
		
		on error resume next
		Set objListaProd = searchProduct(null, text, text, text, null, null, null,arrTargetLang, text, 1, 1, false, bolAddFiles)
		if (objListaProd.Count > 0) then
			bolHasProd = true
		end if		
		if Err.number <> 0 then
			'response.write(Err.description)
			bolHasProd = false
		end if

		if(bolHasNews AND bolHasProd) then
			' sorting result
			Set searchAll = sortDictionary(objListaNews, objListaProd)
		else
			if(bolHasNews)then
				for each x in objListaNews
					searchAll.add x, objListaNews(x)
				next			
			elseif(bolHasProd)then
				for each x in objListaProd
					searchAll.add x, objListaProd(x)
				next			
			end if
		end if

		Set objListaNews = nothing
		Set objListaProd = nothing		
	End Function

	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = doubleValue
		
		'if (Application("dbType") = 0) then
			convertDoubleDelimiter = Replace(convertDoubleDelimiter, ".",",")
		'else		
			'convertDoubleDelimiter = Replace(convertDoubleDelimiter, ",",".")
		'end if			
	End Function
	
	Public Function getTargetCatPageTempl(id_obj, iType)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getTargetCatPageTempl = null		
		
		strSQL = "SELECT categorie.gerarchia, MAX(page_x_template.page_num) AS page_num  FROM target_x_news"
		strSQL = strSQL &" INNER JOIN target ON target_x_news.id_target=target.id"
		strSQL = strSQL &" INNER JOIN target_x_categoria ON target_x_news.id_target=target_x_categoria.id_target"
		strSQL = strSQL &" INNER JOIN categorie ON target_x_categoria.id_categoria=categorie.id"
		strSQL = strSQL &" INNER JOIN page_x_template ON categorie.id_template=page_x_template.id_template"
		strSQL = strSQL &" WHERE target_x_news.id_news=? AND target.type<>3 LIMIT 1;"
		
		strSQL2 = "SELECT categorie.gerarchia, MAX(page_x_template.page_num) AS page_num  FROM target_x_prodotto"
		strSQL2 = strSQL2 &" INNER JOIN target ON target_x_prodotto.id_target=target.id"
		strSQL2 = strSQL2 &" INNER JOIN target_x_categoria ON target_x_prodotto.id_target=target_x_categoria.id_target"
		strSQL2 = strSQL2 &" INNER JOIN categorie ON target_x_categoria.id_categoria=categorie.id"
		strSQL2 = strSQL2 &" INNER JOIN page_x_template ON categorie.id_template=page_x_template.id_template"
		strSQL2 = strSQL2 &" WHERE target_x_prodotto.id_prodotto=? AND target.type<>3 LIMIT 1;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		if(Cint(iType)=1)then
			objCommand.CommandText = strSQL
		else
			objCommand.CommandText = strSQL2		
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_obj)
		Set objRS = objCommand.Execute()	

		if not(objRS.EOF) then
			tmp_ger = objRS("gerarchia")
			tmp_page_num = objRS("page_num")
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			objDict.add "gerarchia", tmp_ger
			objDict.add "page_num", tmp_page_num						
			Set getTargetCatPageTempl = objDict			
			Set objDict = nothing				
		end if

		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			getTargetCatPageTempl = null	
		end if		
	End Function
	
	' implementazione sorting Dictionary per search all tramite algoritmo bubblesort (basse performance al crescere degli elementi ma pratico per il tipo di dati da ordinare)
	Function sortDictionary(aTempDict, bTempDict)
		On Error Resume Next	
		Set sortDictionary = Server.CreateObject("Scripting.Dictionary")

		for each x in aTempDict			
			for each j in bTempDict  
				'response.write("x: "&x&" - aTempDict(x).getTitolo(): "&aTempDict(x).getTitolo()&" --- y: "&y&" - bTempDict(j).getNomeProdotto(): "&bTempDict(j).getNomeProdotto()&"<br>")
				if strComp(aTempDict(x).getTitolo(), bTempDict(j).getNomeProdotto(),1) <= 0 Then
						sortDictionary.add x, aTempDict(x)
						call aTempDict.remove(x)
						exit for
				else
						sortDictionary.add j, bTempDict(j)
						call bTempDict.remove(j)
				end If 

			Next 
			'response.write("<br>")
		Next 
		for each x in aTempDict
			'response.write("x2: "&x&" - aTempDict(x).getTitolo(): "&aTempDict(x).getTitolo()&"<br>")
			sortDictionary.add x, aTempDict(x)
		next
		
		if(Err.number<>0)then
			'response.write(Err.description)
		end if
	End Function
End Class
%>