<%@Language=VBScript codepage=65001 %>
<% 
Response.ContentType="text/html"
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<%
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
	
	Dim id_language, strKeyword, bolDelLanguage, objListaLingue, objDictParams, srtTmpParam, search_key, search_lang_code, items, page,strLangCode,  strValue
	Dim multipleSelection, isMultipleValue, operation, multipleValuesList, arrMultipleValues, arrSingleLineValues, pageRedirect
	
	isMultipleValue = request("is_multiple_selection")
	multipleValuesList = request("multiple_values")
	bolDelLanguage = request("operation")	
	id_language = request("id_multi_language")
	strKeyword = request("keyword")
	'strLangCode = request("lang_code")
	'strValue = request("value")
	'** sostituisco dal titolo:
		'טיאעשל'
	'** con:
		'&egrave;&eacute;&agrave;&ograve;&ugrave;&igrave;&#39;
	strKeyword = Replace(strKeyword, "ט", "&egrave;", 1, -1, 1)
	strKeyword = Replace(strKeyword, "י", "&eacute;", 1, -1, 1)
	strKeyword = Replace(strKeyword, "א", "&agrave;", 1, -1, 1)
	strKeyword = Replace(strKeyword, "ע", "&ograve;", 1, -1, 1)
	strKeyword = Replace(strKeyword, "ש", "&ugrave;", 1, -1, 1)
	strKeyword = Replace(strKeyword, "ל", "&igrave;", 1, -1, 1)
	strKeyword = Replace(strKeyword, "'", "&#39;", 1, -1, 1)
	
	search_key = request("search_key")
	items = request("items")
	page = request("page")
	
	pageRedirect = Application("baseroot")&"/editor/multilanguage/InserisciMultiLingua.asp?search_key="&search_key&"&items="&items&"&page="&page
	
	Dim objLanguage
	Set objLanguage = New LanguageClass


	if(Cint(isMultipleValue) = 0) then
		
		if(Instr(1, typename(objLanguage.getListaLanguageByDesc()), "Dictionary", 1) > 0) then
			Set objListaLanguage = objLanguage.getListaLanguageByDesc()
		
			alreadyModified = false
		
			for each k in objListaLanguage
				srtTmpParam =  request("value_"&k)
				if not(isNull(srtTmpParam)) then
					'** sostituisco dal titolo:
						'טיאעשל'
					'** con:
						'&egrave;&eacute;&agrave;&ograve;&ugrave;&igrave;&#39;
					srtTmpParam = Replace(srtTmpParam, "ט", "&egrave;", 1, -1, 1)
					srtTmpParam = Replace(srtTmpParam, "י", "&eacute;", 1, -1, 1)
					srtTmpParam = Replace(srtTmpParam, "א", "&agrave;", 1, -1, 1)
					srtTmpParam = Replace(srtTmpParam, "ע", "&ograve;", 1, -1, 1)
					srtTmpParam = Replace(srtTmpParam, "ש", "&ugrave;", 1, -1, 1)
					srtTmpParam = Replace(srtTmpParam, "ל", "&igrave;", 1, -1, 1)
					srtTmpParam = Replace(srtTmpParam, "'", "&#39;", 1, -1, 1)	
				end if
				
				if (Cint(id_language) <> -1) then
					if(strComp(bolDelLanguage, "delete", 1) = 0) then
						call objLanguage.deleteMultiLanguage(id_language)
					end if	
				
					if not(alreadyModified) then
						alreadyModified = objLanguage.updateKeywordMultilanguage(id_language,strKeyword)
					end if
				
					if(alreadyModified) then
						call objLanguage.modifyMultiLanguage(strKeyword, k, srtTmpParam)	
					end if
				else
					call objLanguage.insertMultiLanguage(strKeyword, k, srtTmpParam)			
				end if
			next	

			Set objListaLanguage = nothing
		end if

		'*** aggiorno la mappa languageResources
		languageResources.removeAll()
						
		response.Redirect(pageRedirect)	
	elseif(Cint(isMultipleValue) = 1) then
		if (strComp(bolDelLanguage, "delete", 1) = 0) then
			arrMultipleValues = split(multipleValuesList, "|", -1, 1)	
			
			for each xValue in arrMultipleValues
				call objLanguage.deleteMultiLanguage(xValue)				
			next
			
			Set objLanguage = nothing
		else	
			if(Instr(1, typename(objLanguage.getListaLanguageByDesc()), "Dictionary", 1) > 0) then
				Set objListaLanguage = objLanguage.getListaLanguageByDesc()			
				
				arrMultipleValues = split(multipleValuesList, "###", -1, 1)
				
				Dim listSingleVal

				for each xValue in arrMultipleValues
					arrSingleLineValues = split(xValue, "||", -1, 1)
					
					Set listSingleVal = Server.CreateObject("Scripting.Dictionary") 
					alreadyModified = false

					for each yValue in arrSingleLineValues
						if (InStr(1, yValue, "id=", 0) > 0) then
							id_language = Mid(yValue,InStr(1,yValue,"=",0)+1,Len(yValue))
						elseif (InStr(1, yValue, "keyword=", 0) > 0) then 
							'** sostituisco da keyword:
								'טיאעשל'
							'** con:
								'&egrave;&eacute;&agrave;&ograve;&ugrave;&igrave;&#39;
							srtTmpParam = Mid(yValue,InStr(1,yValue,"=",0)+1,Len(yValue))
							srtTmpParam = Replace(srtTmpParam, "ט", "&egrave;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "י", "&eacute;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "א", "&agrave;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "ע", "&ograve;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "ש", "&ugrave;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "ל", "&igrave;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "'", "&#39;", 1, -1, 1)	
							strKeyword = srtTmpParam		
						elseif (InStr(1, yValue, "value_", 0) > 0)  then
							'** sostituisco dal value_:
								'טיאעשל'
							'** con:
								'&egrave;&eacute;&agrave;&ograve;&ugrave;&igrave;&#39;
							srtTmpParam = Mid(yValue,InStr(1,yValue,"=",0)+1,Len(yValue))
							srtTmpParam = Replace(srtTmpParam, "ט", "&egrave;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "י", "&eacute;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "א", "&agrave;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "ע", "&ograve;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "ש", "&ugrave;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "ל", "&igrave;", 1, -1, 1)
							srtTmpParam = Replace(srtTmpParam, "'", "&#39;", 1, -1, 1)							
							strValue = srtTmpParam

							listSingleVal.add Left(yValue,InStr(1,yValue,"=",0)-1),strValue
						end if
					next

					if not(alreadyModified) then
						alreadyModified = objLanguage.updateKeywordMultilanguage(id_language,strKeyword)
					end if
				
					if(alreadyModified) then					
						for each k in objListaLanguage
							call objLanguage.modifyMultiLanguage(strKeyword, k, listSingleVal.Item("value_"&k))
						next
					end if

					Set listSingleVal = nothing		
				next
				Set objListaLanguage = nothing
			end if		
	
			Set objLanguage = nothing
		end if

		'*** aggiorno la mappa languageResources
		languageResources.removeAll()
		
		response.Redirect(pageRedirect)	
	end if

	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>