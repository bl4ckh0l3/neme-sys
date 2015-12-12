<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/SendMailClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->
<!-- #include virtual="/common/include/captcha/adovbs.asp"-->
<!-- #include virtual="/common/include/captcha/iasutil.asp"-->
<!-- #include virtual="/common/include/captcha/functions.asp"-->
<%
'<!--nsys-userproc1-->
%>
<!-- #include virtual="/common/include/Objects/VoucherClass.asp" -->
<%
'<!---nsys-userproc1-->
%>
<%
Response.Buffer = True

' load object
Dim load
Set load = new ImageUploadClass

' calling initialize method
load.initialize

'************* FUNZIONE PER IL VECCHIO CAPTCHA
function TestCaptcha(byval valSession, byval valCaptcha)
	dim tmpSession
	valSession = Trim(valSession)
	valCaptcha = Trim(valCaptcha)
	if (valSession = vbNullString) or (valCaptcha = vbNullString) then
		TestCaptcha = false
	else
		tmpSession = valSession
		valSession = Trim(Session(valSession))
		Session(tmpSession) = vbNullString

		if valSession = vbNullString then
			TestCaptcha = false
		else
			valCaptcha = Replace(valCaptcha,"i","I")
			if StrComp(valSession,valCaptcha,1) = 0 then
				TestCaptcha = true
			else
				TestCaptcha = false
			end if
		end if		
	end if
end function


'** creo la lista di content-type accettati per l'avatar
Set contentTypeList = Server.CreateObject("Scripting.Dictionary")
contentTypeList.add "image/gif","image/gif"
contentTypeList.add "image/jpeg","image/jpeg"
contentTypeList.add "image/jpg","image/jpg"
contentTypeList.add "image/png","image/png"


Dim id_utente, strUserName, strPwd, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, reqTargets, arrTarget, delAvatar
Dim bolUserActive, numSconto, strAdminComments
Dim dateInsertDate, dateModifyDate, objNewsletterUsr
Dim bolPublic, numUserGroup

if(Application("use_recaptcha") = 0) then
	'************* FUNZIONE PER IL VECCHIO CAPTCHA
	' verifico che il codice captcha inserito dall'utente corrisponda con il captcha generato
	' in caso contrario rimando alla pagine di registrazione con errore
	' devo usare campo hidden perchè originale ( load.getValue("captchacode") ) non viene recuperato
	if not TestCaptcha("ASPCAPTCHA",  load.getValue("sent_captchacode")) then
		response.Redirect(Application("baseroot")&"/area_user/manageuser.asp?captcha_err=1")
	end if
else
	'************* RECUPERO PARAMETRI RECAPTCHA
	Dim recaptcha_challenge_field, recaptcha_response_field, recaptcha_private_key, recaptcha_public_key, cTemp
	recaptcha_challenge_field  = load.getValue("sent_recaptcha_challenge_field")
	recaptcha_response_field   = load.getValue("sent_recaptcha_response_field")
	recaptcha_private_key      = Application("recaptcha_priv_key")
	recaptcha_public_key       = Application("recaptcha_pub_key")

	'************* CHECK VALORE RECAPTCHA
	cTemp = recaptcha_confirm(recaptcha_private_key, recaptcha_challenge_field, recaptcha_response_field)
	If cTemp <> "" Then 
		response.Redirect(Application("baseroot")&"/area_user/manageuser.asp?captcha_err=1")
	end if
end if


id_utente = load.getValue("id_utente")
strUserName = load.getValue("username")
strPwd = load.getValue("password")
strEmail = load.getValue("email")
strUsrRuolo = load.getValue("ruolo_utente")
bolPrivacy = load.getValue("privacy")
if(bolPrivacy) then bolPrivacy = 1 else bolPrivacy = 0 end if
bolNewsletter = load.getValue("newsletter")
if(bolNewsletter) then bolNewsletter = 1 else bolNewsletter = 0 end if
bolUserActive = load.getValue("user_active")
numSconto = load.getValue("sconto")
strAdminComments = load.getValue("admin_comments")
dateInsertDate = load.getValue("insertDate")
dateModifyDate = load.getValue("modifyDate")
reqTargets = load.getValue("ListTarget")
arrTarget = split(reqTargets, "|", -1, 1)	
'non funziona, non recupera dalla lista checkbox
'reqNewsletterUsr = load.getValue("list_newsletter")
reqNewsletterUsr = load.getValue("list_newsletter_values")
objNewsletterUsr = split(reqNewsletterUsr, ", ", -1, 1)
delAvatar = load.getValue("del_avatar")
bolPublic = load.getValue("public_profile")
numUserGroup = load.getValue("user_group")
bolAutomaticUser = 0

Dim fileData, fileName, filePath, filePathComplete, fileSize, fileSizeTranslated, contentType, countElements

' File binary data
fileData = load.getFileData("imageupload")
' File name
fileName = LCase(load.getFileName("imageupload"))
' File path
filePath = load.getFilePath("imageupload")
' File path complete
filePathComplete = load.getFilePathComplete("imageupload")
' File size
fileSize = load.getFileSize("imageupload")
' File size translated
fileSizeTranslated = load.getFileSizeTranslated("imageupload")
' Content Type
contentType = load.getContentType("imageupload")
' No. of Form elements
countElements = load.Count

If fileSize > 102400 Then
	response.Redirect(Application("baseroot")&"/area_user/manageuser.asp?error=030")
end if

if(fileSize > 0 AND (Trim(fileName)<>"") AND contentTypeList(contentType)="") then
	response.Redirect(Application("baseroot")&"/area_user/manageuser.asp?error=031")
end if

Dim objUtente, bolDelUtente
Set objUtente = New UserClass
	
Dim objTarget
Set objTarget = New TargetClass

Dim objMail
Set objMail = New SendMailClass

	
Dim objLogger
Set objLogger = New LogClass

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
	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()
	objConn.BeginTrans
		
	call objUtente.modifyUser(id_utente, strUserName, strPwd, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, bolUserActive, numSconto, strAdminComments, dateInsertDate, Now(), bolPublic, numUserGroup, bolAutomaticUser, objConn)	
	
	if(hasUserFields) then
		call objUserField.deleteFieldMatchByUser(id_utente)
		for each k in objListUserField
			On Error Resume Next
				user_field_value = ""
				Set objField = objListUserField(k)
				select Case objField.getTypeField()
				Case 5,6,7
					user_field_value = load.getValue("hidden_"&objUserField.getFieldPrefix()&objField.getID())
				Case Else
					user_field_value = load.getValue(objUserField.getFieldPrefix()&objField.getID())			
				End Select
				
				call objUserField.insertFieldMatch(objField.getID(), id_utente, user_field_value, objConn)			
			if(Err.number<>0) then
				'response.write(Err.description)
			end if
		next
	end if

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
						addTarget=false
						exit for
					end if
				next
			end if
			if(addTarget)then
				call objUtente.insertTargetXUser(j, id_utente, objConn)
			end if
		next
	end if		
	Set listAutomTarget = nothing
	if(Err.number<>0)then
		call objLogger.write("errore inserimento target automatici utente --> id: "&id_utente&"; error: "&Err.description, "system", "error")
	end if

	call objUtente.deleteNewsletterXUser(id_utente, objConn)
	
	if(bolNewsletter) then		
		if not(isNull(objNewsletterUsr)) then
			bolHasAlreadyVoucher = false
			Set objNewsletter = new NewsletterClass
			for each y in objNewsletterUsr
'<!--nsys-userproc2-->
				voucher_campaign = objNewsletter.findNewsletterByID(y).getVoucher()
				if (not(isNull(voucher_campaign)) AND voucher_campaign<>"" AND not(bolHasAlreadyVoucher))then
					Set objVoucher = new VoucherClass
					vcounttmp = objVoucher.countVoucherCodeByCampaign(voucher_campaign, id_utente)
					if (vcounttmp=0)then
						new_voucher_code = objVoucher.generateVoucherCode(voucher_campaign, id_utente, objConn)
						if(new_voucher_code<>"")then
							call objMail.sendMailVoucher(new_voucher_code, 0, strEmail, null, "", lang.getLangCode())
							bolHasAlreadyVoucher = true
						end if
					else
						bolHasAlreadyVoucher = true
					end if
					Set objVoucher = nothing					
				end if
'<!---nsys-userproc2-->
				call objUtente.insertNewsletterXUser(y, id_utente, objConn)
			next
			Set objNewsletter = nothing
		end if
	end if
	
	Dim deletedAv
	deletedAv = false
	if(delAvatar) then
		call objUtente.deleteImageXUser(id_utente, objConn)
		deletedAv = true
	end if
	
	If fileSize > 0 AND (Trim(fileName)<>"") Then
		'Dim id_file
		'id_file = objUtente.findImageIDXUser(id_utente)
		'if not(id_file = "") AND not(isNull(id_file)) then
			'call objUtente.updateImageXUser(id_file, id_utente, fileName, contentType, fileSize, fileData)
		'end if
		
		if not(deletedAv) then
		call objUtente.deleteImageXUser(id_utente, objConn)
		end if
		'******** TODO --> fare inserimento su DB solo se sono immagini ...no file .exe o altro
		call objUtente.insertImageXUser(id_utente, fileName, contentType, fileSize, fileData, objConn)
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
	response.Redirect(Application("baseroot")&"/area_user/confirmRegistration.asp")		
else
	Set objDB = New DBManagerClass
	Set objConn = objDB.openConnection()
	objConn.BeginTrans
	
	Dim idUserMax	
	idUserMax = objUtente.insertUser(strUserName, strPwd, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, bolUserActive, numSconto, strAdminComments, Now(), Now(), bolPublic, numUserGroup, bolAutomaticUser, objConn)
	
	if(hasUserFields) then
		for each k in objListUserField
			On Error Resume Next
				user_field_value = ""
				Set objField = objListUserField(k)
				select Case objField.getTypeField()
				Case 5,6,7
					user_field_value = load.getValue("hidden_"&objUserField.getFieldPrefix()&objField.getID())
				Case Else
					user_field_value = load.getValue(objUserField.getFieldPrefix()&objField.getID())			
				End Select
				call objUserField.insertFieldMatch(objField.getID(), idUserMax, user_field_value, objConn)	
			if(Err.number<>0) then
				'response.write(Err.description)
			end if
		next
	end if
		
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
		'call objLogger.write("errore inserimento target automatici utente --> id: "&idUserMax&"; error: "&Err.description, "system", "error")
	end if
	
	if(bolNewsletter) then
		if not(isNull(objNewsletterUsr)) AND (idUserMax <> -1) then			
			bolHasAlreadyVoucher = false
			Set objNewsletter = new NewsletterClass
			for each y in objNewsletterUsr
'<!--nsys-userproc3-->
				voucher_campaign = objNewsletter.findNewsletterByID(y).getVoucher()
				if (not(isNull(voucher_campaign)) AND voucher_campaign<>"" AND not(bolHasAlreadyVoucher))then
					Set objVoucher = new VoucherClass
					new_voucher_code = objVoucher.generateVoucherCode(voucher_campaign, idUserMax, objConn)
					if(new_voucher_code<>"")then
						call objMail.sendMailVoucher(new_voucher_code, 0, strEmail, null, "", lang.getLangCode())
						bolHasAlreadyVoucher = true
					end if
					Set objVoucher = nothing
				end if
'<!---nsys-userproc3-->
				call objUtente.insertNewsletterXUser(y, idUserMax, objConn)
			next
			Set objNewsletter = nothing
		end if
	end if
	
	If fileSize > 0 AND (Trim(fileName)<>"") Then		
		call objUtente.insertImageXUser(idUserMax, fileName, contentType, fileSize, fileData, objConn)
	end if

	'Spedisco la mail di conferma registrazione
	Dim confirmCode, objUtilTmp
	Set objUtilTmp = new GUIDClass
	if(Application("confirm_registration")=2) then
		confirmCode = objUtilTmp.CreateUserGUID()
		call objUtente.insertConfirmationCodeXUser(idUserMax,confirmCode, objConn)
	end if
	Set objUtilTmp = nothing	
	Set objUtente = nothing
	Set objTarget = nothing	
						
	if objConn.Errors.Count = 0 then
		objConn.CommitTrans
	else
		objConn.RollBackTrans
		response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
	end if			
	Set objDB = nothing
	
	call objMail.sendMailUser(idUserMax, strPwd, strEmail, confirmCode, lang.getLangCode(), 0)
	call objMail.sendMailUser(idUserMax, null, Application("mail_user_receiver"), confirmCode, Application("str_editor_lang_code_default"), 1)
	
	if (Application("confirm_registration")=1) then
		Dim isHTTPS,strLoginAction
		isHTTPS = Request.ServerVariables("HTTPS")
		If isHTTPS = "off" AND Application("use_https") = 1 Then
			strLoginAction = "https://"&Request.ServerVariables("SERVER_NAME")&Application("baseroot")&"/common/include/VerificaUtente.asp"
		Else
			strLoginAction = Application("baseroot")&"/common/include/VerificaUtente.asp"
		End If
		strLoginAction = strLoginAction & "?from=area_user&j_username="&strUserName&"&j_password="&strPwd
		
		response.Redirect(strLoginAction)	
	end if	
		
	response.Redirect(Application("baseroot")&"/area_user/confirmRegistration.asp")				
end if

Set objLogger = nothing
Set objMail = Nothing
Set objUserField =nothing
Set objListUserField = nothing
Set load = nothing
Set objUserLogged = nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>