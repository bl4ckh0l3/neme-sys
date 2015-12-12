<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/SendMailClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->

<%
'On Error Resume Next
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
	
	Dim id_utente, strUserName, strPwd, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, reqTargets, arrTarget
	Dim bolUserActive, numSconto, strAdminComments, bolPublic, numUserGroup
	Dim dateInsertDate, dateModifyDate, objNewsletterUsr


	id_utente = request("id_utente")
	strUserName = request("username")
	strPwd = request("password")
	strEmail = request("email")
	strUsrRuolo = request("ruolo_utente")
	bolPrivacy = request("privacy")
	if(bolPrivacy) then bolPrivacy = 1 else bolPrivacy = 0 end if
	bolNewsletter = request("newsletter")
	if(bolNewsletter) then bolNewsletter = 1 else bolNewsletter = 0 end if
	bolUserActive = request("user_active")
	numSconto = request("sconto")
	if(numSconto="")then
		numSconto=0
	end if
	
	strAdminComments = request("admin_comments")
	bolPublic = request("public_profile")
	numUserGroup = request("user_group")
	dateInsertDate = request("insertDate")
	dateModifyDate = request("modifyDate")
	bolAutomaticUser = 0

	reqTargets = request("ListTarget")
	arrTarget = split(reqTargets, "|", -1, 1)
	
	reqNewsletterUsr = request("list_newsletter")
	objNewsletterUsr = split(reqNewsletterUsr, ",", -1, 1)
	
	Dim objUtente, bolDelUtente
	Set objUtente = New UserClass
	bolDelUtente = request("delete_utente")
	
	Dim objLogger
	Set objLogger = New LogClass
	
	Dim objTarget
	Set objTarget = New TargetClass	
		
	Dim objUserField, objListUserField, hasUserFields
	hasUserFields=false
	On Error Resume Next
	Set objUserField = new UserFieldClass
	Set objListUserField = objUserField.getListUserField(1,"1,3")
	if(objListUserField.count > 0)then
		hasUserFields=true
	end if
	if(Err.number <> 0) then
		hasUserFields=false
	end if

	if (Cint(id_utente) <> -1) then
		if(strComp(bolDelUtente, "del", 1) = 0) then
			Dim canDel
			canDel = objUtente.deleteUser(id_utente)
			if(canDel = true) then
				call objLogger.write("cancellato utente --> id: "&id_utente&"; username: "&strUserName, objUserLogged.getUserName(), "info")
				response.Redirect(Application("baseroot")&"/editor/utenti/ListaUtenti.asp")		
			else
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=029")					
			end if
		
		end if
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		objConn.BeginTrans
			
		call objUtente.modifyUser(id_utente, strUserName, strPwd, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, bolUserActive, numSconto, strAdminComments, dateInsertDate, Now(), bolPublic, numUserGroup, bolAutomaticUser, objConn)
		
		if(hasUserFields) then
			call objUserField.deleteFieldMatchByUser(id_utente)
			for each k in objListUserField
				On Error Resume Next
					Set objField = objListUserField(k)
					select Case objField.getTypeField()
					Case 5,6,7
						user_field_value = request("hidden_"&objUserField.getFieldPrefix()&objField.getID())
					Case Else
						user_field_value = request(objUserField.getFieldPrefix()&objField.getID())			
					End Select
					call objUserField.insertFieldMatch(objField.getID(), id_utente, user_field_value, objConn)			
				if(Err.number<>0) then
					'response.write(Err.description)
				end if
			next
		end if
		call objLogger.write("modificato utente --> id: "&id_utente&"; username: "&strUserName, objUserLogged.getUserName(), "info")
			
		call objUtente.deleteTargetXUser(id_utente, objConn)
				
		if not(isNull(arrTarget)) then
			for each x in arrTarget
				call objUtente.insertTargetXUser(x, id_utente, objConn)
			next
		end if
		
		'***** VERIFICO SE ESISTONO TARGET AUTOMATICI E LI AGGIUNGO A QUELLI SELEZIONATI DALL'UTENTE
		On Error Resume Next
		Set listAutomTarget = objTarget.getListAutomaticTarget()		
		if (Instr(1, typename(listAutomTarget), "Dictionary", 1) > 0) then
			'recupero anche il target per lingua da aggiungere ai target automatici
			Set listLangTarget = objTarget.getListLockedTarget()
			if (Instr(1, typename(listLangTarget), "Dictionary", 1) > 0) then
				for each k in listLangTarget
					listAutomTarget.add k, listLangTarget(k)
				next
			end if
			Set listLangTarget = nothing
			for each j in listAutomTarget
				addTarget=true
				if not(isNull(arrTarget)) then
					for each q in arrTarget
						'call objLogger.write("j: "&j&" q: "&q, "system", "debug")
						if(Cint(j)=Cint(q))then
							'call objLogger.write("j=q: "&j&"-"&q, "system", "debug")
							addTarget=false
							exit for
						end if
					next
				end if
				'call objLogger.write("addTarget: "&addTarget, "system", "debug")
				if(addTarget)then
					call objUtente.insertTargetXUser(j, id_utente, objConn)
				end if
			next
		end if		
		Set listAutomTarget = nothing
		if(Err.number<>0)then
			call objLogger.write("errore inserimento target automatici utente --> id: "&id_utente&"; error: "&Err.description, objUserLogged.getUserName(), "error")
		end if

		call objUtente.deleteNewsletterXUser(id_utente, objConn)
		
		if(bolNewsletter) then		
			if not(isNull(objNewsletterUsr)) then
				for each y in objNewsletterUsr
					call objUtente.insertNewsletterXUser(y, id_utente, objConn)
				next
			end if
		end if
						
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if			
		Set objDB = nothing		
								
		Set objUtente = nothing
		Set objTarget = nothing
		
		response.Redirect(Application("baseroot")&"/editor/utenti/ListaUtenti.asp")		
	else
		Dim idUserMax

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		objConn.BeginTrans
		
		idUserMax =  objUtente.insertUser(strUserName, strPwd, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, bolUserActive, numSconto, strAdminComments, Now(), Now(), bolPublic, numUserGroup, bolAutomaticUser, objConn)
		
		if(hasUserFields) then		
			for each k in objListUserField
				On Error Resume Next
					Set objField = objListUserField(k)
					select Case objField.getTypeField()
					Case 5,6,7
						user_field_value = request("hidden_"&objUserField.getFieldPrefix()&objField.getID())
					Case Else
						user_field_value = request(objUserField.getFieldPrefix()&objField.getID())			
					End Select
					call objUserField.insertFieldMatch(objField.getID(), idUserMax, user_field_value, objConn)	
				if(Err.number<>0) then
					'response.write(Err.description)
				end if
			next
		end if
		call objLogger.write("inserito utente --> id: "&idUserMax&"; username: "&strUserName, objUserLogged.getUserName(), "info")
		
		if not(isNull(arrTarget)) AND (idUserMax <> -1) then
			for each x in arrTarget
				call objUtente.insertTargetXUser(x, idUserMax, objConn)
			next
		end if
		
		'***** VERIFICO SE ESISTONO TARGET AUTOMATICI E LI AGGIUNGO A QUELLI SELEZIONATI DALL'UTENTE
		On Error Resume Next
		Set listAutomTarget = objTarget.getListAutomaticTarget()		
		if (Instr(1, typename(listAutomTarget), "Dictionary", 1) > 0) then
			'recupero anche il target per lingua da aggiungere ai target automatici
			Set listLangTarget = objTarget.getListLockedTarget()
			if (Instr(1, typename(listLangTarget), "Dictionary", 1) > 0) then
				for each k in listLangTarget
					listAutomTarget.add k, listLangTarget(k)
				next
			end if
			Set listLangTarget = nothing
			for each j in listAutomTarget
				addTarget=true
				if not(isNull(arrTarget)) then
					for each q in arrTarget
						'call objLogger.write("j: "&j&" q: "&q, "system", "debug")
						if(Cint(j)=Cint(q))then
							addTarget=false
							exit for
						end if
					next
				end if
				if(addTarget)then
					call objUtente.insertTargetXUser(j, idUserMax, objConn)
				end if
			next
		end if		
		Set listAutomTarget = nothing
		if(Err.number<>0)then
			'call objLogger.write("errore inserimento target automatici utente --> id: "&idUserMax&"; error: "&Err.description, objUserLogged.getUserName(), "error")
		end if
		
		if(bolNewsletter) then
			if not(isNull(objNewsletterUsr)) AND (idUserMax <> -1) then
				for each y in objNewsletterUsr
					call objUtente.insertNewsletterXUser(y, idUserMax, objConn)
				next
			end if
		end if	

		Dim confirmCode, objUtilTmp
		Set objUtilTmp = new UtilClass
		if(Application("confirm_registration")=2) then
			confirmCode = objUtilTmp.CreateUserGUID()
			call objUtente.insertConfirmationCodeXUser(idUserMax,confirmCode, objConn)
		end if
		Set objUtilTmp = nothing	
						
		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if			
		Set objDB = nothing			
		Set objUtente = nothing
		Set objTarget = nothing

		'Spedisco la mail di conferma registrazione
		Dim objMail
		Set objMail = New SendMailClass
		call objMail.sendMailUser(idUserMax, strPwd, strEmail, confirmCode, Application("str_lang_code_default"), 0)
		call objMail.sendMailUser(idUserMax, null, Application("mail_user_receiver"), confirmCode, Application("str_editor_lang_code_default"), 1)
		Set objMail = Nothing

		response.Redirect(Application("baseroot")&"/editor/utenti/ListaUtenti.asp")				
	end if

	Set objUserField =nothing
	Set objListUserField = nothing
	Set objLogger = nothing
	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>