<%

'** imposto il protocollo http o https da usare in tutte le pagine all'occorrenza			
base_url = "http://"
Dim isInitHTTPS
isInitHTTPS = Request.ServerVariables("HTTPS")
If isInitHTTPS = "on" AND Application("use_https") = 1 Then
	base_url = "https://"
End If		


'* questa variabile server per i link attivi nell'editor
Dim cssClass

'/*
 '* impostazioni per il multilingua
 '*/
Dim langEditor

Set langEditor = new LanguageClass
langEditor.setDefaultLangCode(Application("str_editor_lang_code_default"))

'*************** BACKEND ***********

'*** come prima operazione verifico se l'oggetto Application languageResources è valorizzato, e in caso contrario inizializzo la mappa di tutte le chiavi multiingua
if(languageResources.Count <=0) then
  On Error Resume Next
  
  languageResources.removeAll()
  
  Set objListaLanguage = langEditor.getListaLanguageByDesc()

  for each k in objListaLanguage
    languageResources.add k, langEditor.getListaElementsByLang(k)
  next

  Set objListaLanguage = nothing

  if Err.number <> 0 then
    'response.write(Err.description)
  end if  
end if 
  
if not(isNull(Session("langEditorCode"))) AND not isEmpty(Session("langEditorCode")) then
	langEditor.setLangCode(Session("langEditorCode"))  

	'*** controllo se languageResources contiene gli elementi
	if(Instr(1, typename(languageResources.item(langEditor.getLangCode())), "dictionary", 1) > 0)then
 		langEditor.setLangElements(languageResources.item(langEditor.getLangCode()))
	else
      ' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
    	'languageResources.remove(langEditor.getLangCode())
      'languageResources.add langEditor.getLangCode(), langEditor.getListaElementsByLang(langEditor.getDefaultLangCode())
      languageResources.Item(langEditor.getLangCode()) = langEditor.getListaElementsByLang(langEditor.getDefaultLangCode())    
      langEditor.setLangElements(languageResources.item(langEditor.getLangCode()))
	end if
else
	langEditor.setLangCode(Application("str_editor_lang_code_default"))
	Session("langEditorCode") = langEditor.getLangCode() 
  
	'*** controllo se languageResources contiene gli elementi
	if(Instr(1, typename(languageResources.item(langEditor.getLangCode())), "dictionary", 1) > 0)then
 		langEditor.setLangElements(languageResources.item(langEditor.getLangCode()))
	else
      ' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
    	'languageResources.remove(langEditor.getLangCode())
      'languageResources.add langEditor.getLangCode(), langEditor.getListaElementsByLang(langEditor.getDefaultLangCode())
      languageResources.Item(langEditor.getLangCode()) = langEditor.getListaElementsByLang(langEditor.getDefaultLangCode())    
      langEditor.setLangElements(languageResources.item(langEditor.getLangCode()))
	end if
end if

if not(isNull(request("langEditorCode"))) AND not(request("langEditorCode") = "") then
  Dim foundLangList, langElementList
  foundLangList = false

  On Error Resume Next
    call langEditor.setLangCode(request("langEditorCode"))
    Session("langEditorCode") = request("langEditorCode")	 
  
	'*** controllo se languageResources contiene gli elementi
	if(Instr(1, typename(languageResources.item(langEditor.getLangCode())), "dictionary", 1) > 0)then
 		langEditor.setLangElements(languageResources.item(langEditor.getLangCode()))
	else
      ' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
    	'languageResources.remove(langEditor.getLangCode())
      'languageResources.add langEditor.getLangCode(), langEditor.getListaElementsByLang(langEditor.getDefaultLangCode())
      languageResources.Item(langEditor.getLangCode()) = langEditor.getListaElementsByLang(langEditor.getDefaultLangCode())    
      langEditor.setLangElements(languageResources.item(langEditor.getLangCode()))
	end if
  
	if Err.number <> 0 then
		'response.write(Err.description)
	end if  
end if 

'Dim objCache
'Set objCache = New CPageCache
'objCache.AutoCacheToMemory()
'Set objCache = Nothing

'***** VERIFICO SE ESISTE IL COOKIE DELL'UTENTE E LO METTO IN SESSIONE
Dim userCookie
userCookie = request.cookies(Application("srt_default_server_name"))("id_bo")

if (isNull(Session("objCMSUtenteLogged")) OR Session("objCMSUtenteLogged") = "") AND not(isNull(userCookie)) AND not(userCookie="") then
  Dim objUserCookie, objUserCookieTmp, strRuoloCookie
  Set objUserCookieTmp = new UserClass
  Set objUserCookie = objUserCookieTmp.findUserByID(userCookie)
  Set objUserCookieTmp = nothing

  strRuoloCookie = objUserCookie.getRuolo()
  if not(strComp(Cint(strRuoloCookie), Application("guest_role"), 1) = 0) then
    Session("objCMSUtenteLogged") = userCookie
  end if

  Set objUserCookie = nothing  
end if
%>