<%
'** imposto il protocollo http o https da usare in tutte le pagine all'occorrenza			
base_url = "http://"
Dim isInitHTTPS
isInitHTTPS = Request.ServerVariables("HTTPS")
If isInitHTTPS = "on" AND Application("use_https") = 1 Then
	base_url = "https://"
End If		


Dim cssClass

'/*
 '* impostazioni per il multilingua
 '*/
Dim lang

Set lang = new LanguageClass
lang.setDefaultLangCode(Application("str_lang_code_default"))
  
Set objListaLanguage = lang.getListaLanguageByDescExt()

'*************** FRONTEND ***********

'*** come prima operazione verifico se l'oggetto Application languageResources � valorizzato, e in caso contrario inizializzo la mappa di tutte le chiavi multiingua
if(languageResources.Count < lang.countActiveLanguage()) then
  On Error Resume Next
  
  languageResources.removeAll()

  for each k in objListaLanguage
    languageResources.add objListaLanguage(k).getLanguageDescrizione(), lang.getListaElementsByLang(objListaLanguage(k).getLanguageDescrizione())
  next

  if Err.number <> 0 then
    'response.write(Err.description)
  end if  
end if

if not(isNull(Session("langCode"))) AND not isEmpty(Session("langCode")) then
	lang.setLangCode(Session("langCode"))

	'*** controllo se languageResources contiene gli elementi
	if(Instr(1, typename(languageResources.item(lang.getLangCode())), "dictionary", 1) > 0)then
 		lang.setLangElements(languageResources.item(lang.getLangCode()))
	else
    languageResources.Item(lang.getLangCode()) = lang.getListaElementsByLang(lang.getDefaultLangCode())
    lang.setLangElements(languageResources.item(lang.getLangCode()))
	end if
else
	if(Application("use_locale") = "1") then
		'**********tento di recuperare il Locale dell'utente, se non trovo corrispondenze sulla mappa imposto la lingua di default del CMS
		Dim localeLang, mappedLocale
		localeLang = GetLocale()
		mappedLocale = ""
		
		if not(isNull(objListPairKeyValueLocale)) AND (Instr(1, typename(objListPairKeyValueLocale), "dictionary", 1) > 0) then
			mappedLocale = objListPairKeyValueLocale.Item(CStr(localeLang))
		end if
		
		if not(mappedLocale = "") then
			lang.setLangCode(mappedLocale)
		else
			lang.setLangCode(Application("str_lang_code_default"))
		end if
	else
		lang.setLangCode(Application("str_lang_code_default"))
	end if
	
	Session("langCode") = lang.getLangCode()
  
	'*** controllo se languageResources contiene gli elementi
	if(Instr(1, typename(languageResources.item(lang.getLangCode())), "dictionary", 1) > 0)then
 		lang.setLangElements(languageResources.item(lang.getLangCode()))
	else
    languageResources.Item(lang.getLangCode()) = lang.getListaElementsByLang(lang.getDefaultLangCode())
		lang.setLangElements(languageResources.item(lang.getLangCode()))
	end if
end if

if not(isNull(request("langCode"))) AND not(request("langCode") = "") then
  Dim foundLangList, langElementList
  foundLangList = false

  On Error Resume Next
  lang.setLangCode(request("langCode"))
  Session("langCode") = request("langCode")	  
  
	'*** controllo se languageResources contiene gli elementi
	if(Instr(1, typename(languageResources.item(lang.getLangCode())), "dictionary", 1) > 0)then
 		lang.setLangElements(languageResources.item(lang.getLangCode()))
	else
    languageResources.Item(lang.getLangCode()) = lang.getListaElementsByLang(lang.getDefaultLangCode())
    lang.setLangElements(languageResources.item(lang.getLangCode()))
	end if
  
	if Err.number <> 0 then
		'response.write(Err.description)
	end if 
else
  '************   SE NON VIENE PASSATA LA LINGUA DALLA URL VERIFICO SE NELLA URL DELLA PAGINA È CONTENUTO IL CODICE LINGUA
  '************   SE È UNA LINGUA ATTIVA E SE È DIVERSA DALLA LINGUA IN SESSIONE E IN TAL CASO IMPOSTO LA LINGUA RECUPERATA
  On Error Resume Next
  Dim retrivedLangURL, forcedLangCode

  forcedLangCode = lang.getLangCode()
  retrivedLangURL = Request.ServerVariables("URL")

  for each r in objListaLanguage
    if (objListaLanguage(r).isLangActive()="1") then
      if(Instr(1, retrivedLangURL, "/"&objListaLanguage(r).getLanguageDescrizione()&"/", 1) > 0) then
        forcedLangCode = objListaLanguage(r).getLanguageDescrizione()
        exit for
      end if
    end if
  next

  lang.setLangCode(forcedLangCode)
  Session("langCode") = forcedLangCode	
  
	'*** controllo se languageResources contiene gli elementi
	if(Instr(1, typename(languageResources.item(lang.getLangCode())), "dictionary", 1) > 0)then
 		lang.setLangElements(languageResources.item(lang.getLangCode()))
	else
    languageResources.Item(lang.getLangCode()) = lang.getListaElementsByLang(lang.getDefaultLangCode())
    lang.setLangElements(languageResources.item(lang.getLangCode()))
	end if
  
	if Err.number <> 0 then
		'response.write(Err.description)
	end if   
end if
  
Set objListaLanguage = nothing

'***** VERIFICO SE ESISTE IL COOKIE DELL'UTENTE E LO METTO IN SESSIONE
Dim userCookie
userCookie = request.cookies(Application("srt_default_server_name"))("id_user")

if (isNull(Session("objUtenteLogged")) OR Session("objUtenteLogged") = "") AND not(isNull(userCookie)) AND not(userCookie="") then
  Dim objUserCookie, objUserCookieTmp, strRuoloCookie
  Set objUserCookieTmp = new UserClass
  Set objUserCookie = objUserCookieTmp.findUserByID(userCookie)
  Set objUserCookieTmp = nothing

  strRuoloCookie = objUserCookie.getRuolo()
  if not(strComp(Cint(strRuoloCookie), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloCookie), Application("editor_role"), 1) = 0) then
    Session("objUtenteLogged") = userCookie
	Session("objUtenteOnline") = objUserCookie.getUserID()&"|"&objUserCookie.getPublic()&"|"&objUserCookie.hasImageUser(objUserCookie.getUserID())&"|"&objUserCookie.getUserName()
	
	On Error Resume Next
	'**** aggiungo l'utent eloggato alla lista degli utenti online
	for each x in onlineUsersList			
		if onlineUsersList(x)=Session("objUtenteOnline") then
			onlineUsersList.remove(x)
		end if
	next
  onlineUsersList.Item(Session.SessionID) = Session("objUtenteOnline")

	'response.cookies(Application("srt_default_server_name"))("user_online")=Session("objUtenteOnline")
	'response.cookies(Application("srt_default_server_name")).Expires=DateAdd("m",6,date())
		
	if(Err.number <>0) then			
	end if

  end if

  Set objUserCookie = nothing  
end if
%>
