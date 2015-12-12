<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->

<%
On error resume next

'** procedure per la gestione dei file e directory ti tutti i template in base alla lingua selezionata
Public Sub createLangDirTemplate(langCode)
	Dim fso, folderTemplate,fold
	
	'** preparo l'oggetto FileSystemObject e i path necessari per creare o cancellare i files e le directory dei template legati alla lingua selezionata
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	folderTemplate = server.mappath(Application("baseroot")&Application("dir_upload_templ"))
	langCode=Ucase(langCode)

	Set fold = fso.GetFolder(folderTemplate) 
	
	Dim subfold, subfoldpath,subfoldpathnew
		
	for each subfolder in fold.subFolders
		subfold=subfolder.Name
		subfoldpath=fso.BuildPath(folderTemplate,subfold)
		
		subfoldpathnew=fso.BuildPath(subfoldpath,langCode)
		
		if not(fso.FolderExists(subfoldpathnew)) then
			fso.CreateFolder(subfoldpathnew)
		end if
		
		'** copio i file dalla dir principale alla nuova sottodir della lingua specificata
		fso.CopyFile subfoldpath&"\*.asp",subfoldpathnew&"\"
		'** copio tutta la dir include nella nuova sottodir della lingua specificata
		if (fso.FolderExists(fso.BuildPath(subfoldpath,"include"))) then
			fso.CopyFolder fso.BuildPath(subfoldpath,"include"),subfoldpathnew&"\"
		end if
		
		'call objLogger.write("inserisco directory template per attivazione lingua--> folder: "&folderTemplate&"; langCode: "&langCode&"; subfold: "&subfold&"; subfoldpath: "&subfoldpath&"; subfoldpathnew: "&subfoldpathnew, "system", "debug")
	next 
	
	Set fold = Nothing 	
	
	' elimino il FileSystemObject
	Set fso=Nothing	
End Sub

Public Sub deleteLangDirTemplate(langCode)
	Dim fso, folderTemplate,fold
	
	'** preparo l'oggetto FileSystemObject e i path necessari per creare o cancellare i files e le directory dei template legati alla lingua selezionata
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	folderTemplate = server.mappath(Application("baseroot")&Application("dir_upload_templ"))
	langCode=Ucase(langCode)

	Set fold = fso.GetFolder(folderTemplate) 
	
	Dim subfold, subfoldpath,subfoldpathnew
	
	for each subfolder in fold.subFolders
		subfold=subfolder.Name
		subfoldpath=fso.BuildPath(folderTemplate,subfold)
		
		subfoldpathnew=fso.BuildPath(subfoldpath,langCode)
		
		if (fso.FolderExists(subfoldpathnew)) then
			fso.DeleteFolder(subfoldpathnew)
		end if
		
		'call objLogger.write("cancello directory template per disattivazione lingua--> folder: "&folderTemplate&"; subfold: "&subfold&"; subfoldpath: "&subfoldpath&"; subfoldpathnew: "&subfoldpathnew, "system", "debug")
	next 
	
	Set fold = Nothing 	
	
	' elimino il FileSystemObject
	Set fso=Nothing	
End Sub

if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Set objLogger = New LogClass
	
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if
	
	Dim id_language, strDescrizione, bolDelLanguage, strLabel, set_target_to_users, subdomain_active, lang_active
	id_language = request("id_language")
	strDescrizione = request("descrizione")
	strLabel = request("selected_label")
	bolDelLanguage = request("delete_language")
	set_target_to_users = request("set_target_to_users")		
	lang_active = request("lang_active")
	subdomain_active = request("subdomain_active")
	url_subdomain = request("url_subdomain")
	
	Dim objLanguage, objTarget, objTmpTarget
	Set objLanguage = New LanguageClass
	Set objTarget = New TargetClass
	
	if (Cint(id_language) <> -1) then
		if(strComp(bolDelLanguage, "del", 1) = 0) then
			call objLanguage.deleteLanguage(id_language)
			Set objTmpTarget = objTarget.findTargetByDescNoTransaction(Application("strLangPrefix")&strDescrizione)
			call objTarget.modifyTargetNoTransaction(objTmpTarget.getTargetID(), objTmpTarget.getTargetDescrizione(), objTmpTarget.getTargetType(), 0, 0)
			Set objTmpTarget = nothing				
			Set objLanguage = nothing
			Set objTarget = nothing
			
			'** cancello la directory della lingua selezionata dentro tutti i template caricati
			call deleteLangDirTemplate(strDescrizione)
			Set objLogger = nothing
			response.Redirect(Application("baseroot")&"/editor/language/InserisciLingua.asp")			
		end if		
	else
		call objLanguage.insertLanguage(strDescrizione, strLabel, lang_active, subdomain_active, url_subdomain)
		
		if not(isNull(objTarget.findTargetByDescNoTransaction(Application("strLangPrefix")&strDescrizione))) then
			Set objTmpTarget = objTarget.findTargetByDescNoTransaction(Application("strLangPrefix")&strDescrizione)
			call objTarget.modifyTargetNoTransaction(objTmpTarget.getTargetID(), objTmpTarget.getTargetDescrizione(), objTmpTarget.getTargetType(), 1, 0)
			Set objTmpTarget = nothing
		else
			call objTarget.insertTargetNoTransaction(Application("strLangPrefix")&strDescrizione, 3, 1, 0)
		end if
		
		Dim objUsertmp, objListUsrTmp, usrTmp, objTargetTmp
		if(set_target_to_users = "1") then
			Set objUsertmp = new UserClass
			Set objListUsrTmp = objUsertmp.findUtente(null, null, 1, null, 0, null)
			Set objTargetTmp = objTarget.findTargetByDescNoTransaction(Application("strLangPrefix")&strDescrizione)
			
			for each x in objListUsrTmp
				Set usrTmp = objListUsrTmp(x)
				'if(usrTmp.getRuolo() = Application("admin_role") OR usrTmp.getRuolo() = Application("editor_role")) then
					call objUsertmp.insertTargetXUserNoTransaction(objTargetTmp.getTargetID(), usrTmp.getUserID())	
				'end if			
			next
			
			Set usrTmp = nothing
			Set objTargetTmp = nothing
			Set objListUsrTmp = nothing
			Set objUsertmp = nothing
		end if
			
		'** creo la directory della lingua selezionata dentro tutti i template caricati
		call createLangDirTemplate(strDescrizione)
		
		Set objLanguage = nothing
		Set objTarget = nothing
		Set objLogger = nothing
		response.Redirect(Application("baseroot")&"/editor/language/InserisciLingua.asp")				
	end if

	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>