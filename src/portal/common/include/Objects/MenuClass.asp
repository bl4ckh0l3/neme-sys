<%
Class MenuClass

	Public Function getCompleteMenu()
		getCompleteMenu = null
		
		Dim objCompleteMenu, objCategoria, objListaCats, completeMenu
		Set completeMenu = Server.CreateObject("Scripting.Dictionary")
		Set objCategoria = new CategoryClass
		Set objListaCats = objCategoria.getListaCategorie()
		if (not isNull(objListaCats) AND not isEmpty(objListaCats)) then
			Dim tmpObjCat, x
			for each x in objListaCats
				Set tmpObjCat = objListaCats(x)
				if(tmpObjCat.isCatVisible()) then
					completeMenu.add tmpObjCat.getCatGerarchia(), tmpObjCat
				end if
			next
			
			Set getCompleteMenu = completeMenu
		end if
		
		Set objListaCats = nothing
		Set objCategoria = nothing
		Set completeMenu = nothing
	End Function

	Public Function getCompleteMenuByMenu(iNumMenu)
		getCompleteMenuByMenu = null
		
		Dim objCompleteMenu, objCategoria, objListaCats, completeMenu
		Set completeMenu = Server.CreateObject("Scripting.Dictionary")
		Set objCategoria = new CategoryClass
		On Error Resume Next
		Set objListaCats = objCategoria.getListaCategorieByMenu(iNumMenu)
		if (not isNull(objListaCats) AND not isEmpty(objListaCats)) then
			Dim tmpObjCat, x
			for each x in objListaCats
				Set tmpObjCat = objListaCats(x)
				if(tmpObjCat.isCatVisible()) then
					completeMenu.add tmpObjCat.getCatGerarchia(), tmpObjCat
				end if
			next
			
			Set getCompleteMenuByMenu = completeMenu
		end if
		if(Err.number <>0) then
			getCompleteMenuByMenu = null
		end if
		
		Set objListaCats = nothing
		Set objCategoria = nothing
		Set completeMenu = nothing
	End Function

	Public Function getRangeMenu(strGerDa, strGerA, iLevel, bolOnlyThisLev)
		getRangeMenu = null
		strGerDa = gerarchia2double(strGerDa)
		strGerA = gerarchia2double(strGerA)
		
		Dim objCompleteMenu, objCategoria, objListaCats
		Set completeMenu = Server.CreateObject("Scripting.Dictionary")
		Set objCategoria = new CategoryClass
		Set objListaCats = objCategoria.getListaCategorie()
		if (not isNull(objListaCats) AND not isEmpty(objListaCats)) then
			Dim tmpObjCat, strGerTmp
			for each x in objListaCats
				Set tmpObjCat = objListaCats(x)
				if(tmpObjCat.isCatVisible()) then
					if(Cbool(bolOnlyThisLev)) then					
						if (getLivello(tmpObjCat.getCatGerarchia()) = Cint(iLevel)) then
							strGerTmp = gerarchia2double(tmpObjCat.getCatGerarchia())
							if(strGerTmp >= strGerDa AND strGerTmp <= strGerA) then
								completeMenu.add tmpObjCat.getCatGerarchia(), tmpObjCat
							end if
						end if						
					else
						if (getLivello(tmpObjCat.getCatGerarchia()) >= Cint(iLevel)) then
							strGerTmp = gerarchia2double(tmpObjCat.getCatGerarchia())
							if(strGerTmp >= strGerDa AND strGerTmp <= strGerA) then
								completeMenu.add tmpObjCat.getCatGerarchia(), tmpObjCat
							end if
						end if
					end if							
				end if
			next
			
			Set getRangeMenu = completeMenu
		end if
		
		Set objListaCats = nothing
		Set objCategoria = nothing
		Set completeMenu = nothing
	End Function
	
	Public Function getLivello(iGerarchia)
		getLivello = calculateLevel(iGerarchia)
	End Function
	
    Private Function  calculateLevel(iGerarchia)
        Dim strSeparator, arrTmp
		strSeparator = "."

		calculateLevel = 1
		
		if not(isNull(iGerarchia)) AND not(iGerarchia = "") then
			arrTmp = Split(iGerarchia, strSeparator, -1, 1)
			calculateLevel = calculateLevel + UBound(arrTmp)
		end if
    End Function
	
	

     ' Converte la gerarchia in un numero in virgola mobile direttamente comparabile nella forma 0.030201.
     '
     ' @param gerarchia la gerarchia da convertire
     ' @return il valore in virgola mobile calcolato

    Public Function gerarchia2double(strgerarchia)
        Dim gerarchiaDbl, scale, begin, end_, level, levelInt
		gerarchiaDbl= 0.0
        scale= 1.0 / 100.0
        begin= 1

        do 
            level= Mid(strgerarchia, begin, 2)			
            levelInt= Cint(level)			
            gerarchiaDbl = gerarchiaDbl + (Cdbl(levelInt)) * scale			
            scale = scale / 100.0
			begin= begin + 3
        loop While begin < Len(strgerarchia)

        gerarchia2double =  gerarchiaDbl			
    End Function
	
	Public Function getTipsMenu(strGerarchia)		
		Dim objCategoria, objCatTmp, menuTipsDic, strTmp
		Set menuTipsDic = Server.CreateObject("Scripting.Dictionary")
		Set objCategoria = new CategoryClass
		
		strTmp = strGerarchia
		
		getTipsMenu = null
		Dim arrCounter, counter
		arrCounter = Split(strTmp, ".", -1, 1)
		
		strTmp = ""
		for counter = 0 to UBound(arrCounter)
			strTmp = strTmp & arrCounter(counter)
			Set objCatTmp = objCategoria.findCategoriaByGerarchia(strTmp)
			menuTipsDic.add objCatTmp.getCatGerarchia(), objCatTmp
			strTmp = strTmp & "."
			Set objCatTmp = nothing
		next
		Set getTipsMenu = menuTipsDic	
		
		Set objCategoria = nothing
		Set menuTipsDic = nothing	
	End Function
	
	Public Function resolveHrefUrl(strBaseUrl, modPageNum, objLang, objCatSelected, objTemplateSelected, objPage4Template)
		resolveHrefUrl = "#"
		tmpBaseUrl = strBaseUrl

		On Error Resume Next
		'*** recupero la dir language o il sottodominio per lingua
		langCode = UCase(objLang.getLangCode())
		langcodeDir = langCode & "/"
		isLangSubdomainActive = objLang.isLanguageSelectedSubDomainActive()
		if(isLangSubdomainActive) then langcodeDir = ""  end if
		
		'*** verifico esistenza del sottodominio per categoria e compongo la base_url
		catSubDomUrl = Replace(objCatSelected.getSubDomainURL(), " ", "", 1, -1, 1)

		if not(catSubDomUrl = "")then
			tmpBaseUrl = tmpBaseUrl & catSubDomUrl
		else
			if (isLangSubdomainActive)then
				tmpBaseUrl = tmpBaseUrl & objLang.getCurrURLSubDomainByLangCode(langCode)
			else
				tmpBaseUrl = tmpBaseUrl & Application("srt_default_server_name")
			end if
		end if
		
		templateDir = objTemplateSelected.getDirTemplate()
		templatePage = objPage4Template.findPageByNum(objTemplateSelected.getID(), modPageNum)	
		
		'************** VERIFICO SE DEVE ESSERE USATO URL REWRITE
		if(Application("use_url_rewrite")=1) then
			'/IT/detail/the-product
			templatePage = Replace(templatePage, ".asp", "", 1, -1, 1)
			resolveHrefUrl = tmpBaseUrl & Application("baseroot") & "/" & langcodeDir & templatePage & "/" & templateDir
		else
			resolveHrefUrl = tmpBaseUrl & Application("baseroot") & Application("dir_upload_templ")& templateDir & "/" & langcodeDir & templatePage	
		end if
		
		if(Err.number <> 0) then
			resolveHrefUrl = "#"
		end if	
	End Function
End Class
%>