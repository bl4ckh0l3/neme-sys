<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Type Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->

<%
Class SendMailClass

	Public Sub sendMailUser(id_user, strPassword, strEmail, confirmCode, langCode, isAdmin)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if

		
		'******************************************************************************
		'		Predispongo la lista di indirizzi email per l'invio della conferma registrazione;
		'		Gli indirizzi si distinguono per zona geografica, in base al campo
		'		"Country" selezionato dell'utente viene scelta una mail o un'altra
		'******************************************************************************
		Dim objListaMail, strMailCc, strMailBcc
		strMailCc = Application("mail_user_cc")
		strMailBcc = Application("mail_user_bcc")
		
		'Set objListaMail = Server.CreateObject("Scripting.Dictionary")				
		'objListaMail.add lang.getTranslated("Grecia"), Application("mail_cc_southern")
		'objListaMail.add lang.getTranslated("Italia"), Application("mail_cc_southern")																																														
		'strMailCc = objListaMail.Item(strCountry)		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
													
		With iMsg
			Set .Configuration = iConf
				.To = strEmail
				.From = Application("mail_user_sender")
				.Sender = Application("mail_user_sender")
				if(Trim(strMailCc) <> "") then .Cc = strMailCc end if
				if(Trim(strMailBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.utenti.mail.subject.label.confirm_registration")
				if(CBool(isAdmin)) then
					.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/utenti/mail_notif/mail_utente_admin.asp?id_utente="&id_user&"&lang_code="&langCode&"&confirm_code="&confirmCode&"&sid="&Rnd(session.SessionID), 31
				else
					.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/utenti/mail_notif/mail_utente_guest.asp?id_utente="&id_user&"&password="&strPassword&"&lang_code="&langCode&"&confirm_code="&confirmCode&"&sid="&Rnd(session.SessionID), 0
				end if
				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub

	Public Sub sendMailUserPwd(id_user, strPassword, strEmail, langCode)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		'******************************************************************************
		'		Predispongo la lista di indirizzi email per l'invio della conferma registrazione;
		'		Gli indirizzi si distinguono per zona geografica, in base al campo
		'		"Country" selezionato dell'utente viene scelta una mail o un'altra
		'******************************************************************************
		Dim objListaMail, strMailCc, strMailBcc
		strMailCc = Application("mail_user_cc")
		strMailBcc = Application("mail_user_bcc")
		
		'Set objListaMail = Server.CreateObject("Scripting.Dictionary")				
		'objListaMail.add lang.getTranslated("Grecia"), Application("mail_cc_southern")
		'objListaMail.add lang.getTranslated("Italia"), Application("mail_cc_southern")																																														
		'strMailCc = objListaMail.Item(strCountry)		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
													
		With iMsg
			Set .Configuration = iConf
				.To = strEmail
				.From = Application("mail_user_sender")
				.Sender = Application("mail_user_sender")
				if(Trim(strMailCc) <> "") then .Cc = strMailCc end if
				if(Trim(strMailBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.utenti.mail.subject.label.confirm_new_password")
				.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/utenti/mail_notif/mail_pwd_guest.asp?id_utente="&id_user&"&password="&strPassword&"&lang_code="&langCode&"&sid="&Rnd(session.SessionID), 0

				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub
	
	Public Sub sendMailOrder(IDOrder, strEmail, isAdmin, langCode)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		'******************************************************************************
		'		Predispongo la lista di indirizzi email per l'invio della conferma registrazione;
		'		Gli indirizzi si distinguono per zona geografica, in base al campo
		'		"Country" selezionato dell'utente viene scelta una mail o un'altra
		'******************************************************************************
		Dim objListaMail, strMailOrderCc, strMailOrderBcc
		strMailOrderCc = Application("mail_order_cc")
		strMailOrderBcc = Application("mail_order_bcc")
		
		'Set objListaMail = Server.CreateObject("Scripting.Dictionary")				
		'objListaMail.add lang.getTranslated("Grecia"), Application("mail_cc_southern")
		'objListaMail.add lang.getTranslated("Italia"), Application("mail_cc_southern")																																														
		'strMailCc = objListaMail.Item(strCountry)		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update

		With iMsg
			Set .Configuration = iConf
				.To = strEmail
				.From = Application("mail_order_sender")
				.Sender = Application("mail_order_sender")
				if(Trim(strMailOrderCc) <> "") then .Cc = strMailCc end if
				if(Trim(strMailOrderBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.utenti.mail.subject.label.confirm_order")
				if(CBool(isAdmin)) then
					.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/ordini/mail_notif/mail_ordine_admin.asp?lang_code="&langCode&"&id_ordine="&IDOrder&"&sid="&Rnd(session.SessionID), 31
				else
					.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/ordini/mail_notif/mail_ordine_guest.asp?lang_code="&langCode&"&id_ordine="&IDOrder&"&sid="&Rnd(session.SessionID), 0
				end if
				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub
	
	
	Public Sub sendMailOrderDown(IDOrder, strEmail, isAdmin, langCode)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		'******************************************************************************
		'		Predispongo la lista di indirizzi email per l'invio della conferma registrazione;
		'		Gli indirizzi si distinguono per zona geografica, in base al campo
		'		"Country" selezionato dell'utente viene scelta una mail o un'altra
		'******************************************************************************
		Dim objListaMail, strMailOrderCc, strMailOrderBcc
		strMailOrderCc = Application("mail_order_cc")
		strMailOrderBcc = Application("mail_order_bcc")
		
		'Set objListaMail = Server.CreateObject("Scripting.Dictionary")				
		'objListaMail.add lang.getTranslated("Grecia"), Application("mail_cc_southern")
		'objListaMail.add lang.getTranslated("Italia"), Application("mail_cc_southern")																																														
		'strMailCc = objListaMail.Item(strCountry)		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update

		With iMsg
			Set .Configuration = iConf
				.To = strEmail
				.From = Application("mail_order_sender")
				.Sender = Application("mail_order_sender")
				if(Trim(strMailOrderCc) <> "") then .Cc = strMailCc end if
				if(Trim(strMailOrderBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.utenti.mail.subject.label.confirm_down_order")
				if(CBool(isAdmin)) then
					.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/ordini/mail_notif/mail_down_ordine_admin.asp?lang_code="&langCode&"&id_ordine="&IDOrder&"&sid="&Rnd(session.SessionID), 31
				else
					.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/ordini/mail_notif/mail_down_ordine_guest.asp?lang_code="&langCode&"&id_ordine="&IDOrder&"&sid="&Rnd(session.SessionID), 0
				end if
				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub
	
	Public Sub sendMailProdEndDisp(IDProd, strEmail, isAdmin, langCode)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		'******************************************************************************
		'		Predispongo la lista di indirizzi email per l'invio della conferma registrazione;
		'		Gli indirizzi si distinguono per zona geografica, in base al campo
		'		"Country" selezionato dell'utente viene scelta una mail o un'altra
		'******************************************************************************
		Dim objListaMail, strMailOrderCc, strMailOrderBcc
		strMailOrderCc = Application("mail_order_cc")
		strMailOrderBcc = Application("mail_order_bcc")
		
		'Set objListaMail = Server.CreateObject("Scripting.Dictionary")				
		'objListaMail.add lang.getTranslated("Grecia"), Application("mail_cc_southern")
		'objListaMail.add lang.getTranslated("Italia"), Application("mail_cc_southern")																																														
		'strMailCc = objListaMail.Item(strCountry)		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update

		With iMsg
			Set .Configuration = iConf
				.To = strEmail
				.From = Application("mail_order_sender")
				.Sender = Application("mail_order_sender")
				if(Trim(strMailOrderCc) <> "") then .Cc = strMailCc end if
				if(Trim(strMailOrderBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.utenti.mail.subject.label.prod_end")
				if(CBool(isAdmin)) then
					.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/prodotti/mail_notif/mail_prodotto_no_disp.asp?lang_code="&langCode&"&id_prodotto="&IDProd&"&sid="&Rnd(session.SessionID), 31
				else
					.CreateMHTMLBody strServerName&Application("baseroot") & "/editor/prodotti/mail_notif/mail_prodotto_no_disp.asp?lang_code="&langCode&"&id_prodotto="&IDProd&"&sid="&Rnd(session.SessionID), 0
				end if
				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub

	Public Sub sendMailContactus(mailTo, userMail, mailText, nome, cognome, telefono, indirizzo, zipcode, citta, nazione, templatepath, langCode)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		'******************************************************************************
		'		Predispongo la lista di indirizzi email per l'invio della conferma;
		'		Gli indirizzi si distinguono per zona geografica, in base al campo
		'		"Country" selezionato dell'utente viene scelta una mail o un'altra
		'******************************************************************************
		Dim objListaMail, strMailCc, strMailBcc
		strMailCc = Application("mail_cc")
		strMailBcc = Application("mail_bcc")
		
		'Set objListaMail = Server.CreateObject("Scripting.Dictionary")				
		'objListaMail.add lang.getTranslated("Grecia"), Application("mail_cc_southern")
		'objListaMail.add lang.getTranslated("Italia"), Application("mail_cc_southern")																																														
		'strMailCc = objListaMail.Item(strCountry)		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
		
		With iMsg
			Set .Configuration = iConf
				.To = Application("mail_receiver")
				.From = Application("mail_sender")
				.Sender = Application("mail_sender")
				if(Trim(strMailCc) <> "") then .Cc = strMailCc end if
				if(Trim(strMailBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.utenti.mail.subject.label.user_contact")
				.CreateMHTMLBody strServerName&Application("baseroot")&Application("dir_upload_templ")&templatepath&"?lang_code="&langCode&"&userMail="&userMail&"&mailText="&mailText&"&nome="&nome&"&cognome="&cognome&"&telefono="&telefono&"&indirizzo="&indirizzo&"&zipcode="&zipcode&"&citta="&citta&"&nazione="&nazione&"&sid="&Rnd(session.SessionID), 0				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub

	Public Sub sendMailModules(mailTo, userMail, mailText, telefono, tmpFilePath, templatepath, langCode)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		'******************************************************************************
		'		Predispongo la lista di indirizzi email per l'invio della conferma;
		'		Gli indirizzi si distinguono per zona geografica, in base al campo
		'		"Country" selezionato dell'utente viene scelta una mail o un'altra
		'******************************************************************************
		Dim objListaMail, strMailCc, strMailBcc
		strMailCc = Application("mail_cc")
		strMailBcc = Application("mail_bcc")
		
		'Set objListaMail = Server.CreateObject("Scripting.Dictionary")				
		'objListaMail.add lang.getTranslated("Grecia"), Application("mail_cc_southern")
		'objListaMail.add lang.getTranslated("Italia"), Application("mail_cc_southern")																																														
		'strMailCc = objListaMail.Item(strCountry)		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
		
		With iMsg
			Set .Configuration = iConf
				.To = Application("mail_receiver")
				.From = Application("mail_sender")
				.Sender = Application("mail_sender")
				if(Trim(strMailCc) <> "") then .Cc = strMailCc end if
				if(Trim(strMailBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.utenti.mail.subject.label.receive_module")				
				.CreateMHTMLBody strServerName&Application("baseroot")&Application("dir_upload_templ")&templatepath&"?lang_code="&langCode&"&userMail="&userMail&"&mailText="&Server.URLEncode(mailText)&"&telefono="&telefono&"&filepath="&Server.URLEncode(tmpFilePath)&"&sid="&Rnd(session.SessionID), 0					
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub
	

	Public Sub sendMailTellaFriend(userMail, mailTo, pageURL, tellafriendMessage, langCode)
		on error resume next	

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
		
		With iMsg
			Set .Configuration = iConf
				.To = mailTo
				.From = userMail
				.Sender = Application("mail_sender")
				.Subject = langMail.getTranslated("backend.utenti.mail.subject.label.tella_friend")
				.CreateMHTMLBody strServerName&Application("baseroot")&"/common/include/mail_tellafriend.asp?lang_code="&langCode&"&userMail="&userMail&"&pageURL="&pageURL&"&tellafriendMessage="&tellafriendMessage&"&sid="&Rnd(session.SessionID), 0
				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub
	

	Public Sub sendMailComment(idComment, mailTo, langCode)
		on error resume next	

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
		
		With iMsg
			Set .Configuration = iConf
				.To = mailTo
				.From = mailTo
				.Sender = Application("mail_sender")
				.Subject = langMail.getTranslated("backend.comments.mail.subject.label.confirm_comment")
				.CreateMHTMLBody strServerName&Application("baseroot")&"/common/include/mail_comment.asp?lang_code="&langCode&"&idComment="&idComment&"&sid="&Rnd(session.SessionID), 0
				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub
	

	Public Sub sendMailCheckFriend(idFriend, mailTo, id_utente, active, langCode, action)
		on error resume next	

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
		
		With iMsg
			Set .Configuration = iConf
				.To = mailTo
				.From = Application("mail_sender")
				.Sender = Application("mail_sender")
				.Subject = langMail.getTranslated("backend.comments.mail.subject.label.check_friend")
				.CreateMHTMLBody strServerName&Application("baseroot")&"/area_user/mail_checkfriend.asp?lang_code="&langCode&"&idFriend="&idFriend&"&idUtente="&id_utente&"&action="&action&"&active="&active&"&sid="&Rnd(session.SessionID), 0
				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub

	Public Sub sendMailAds(mailTo, mailText, id_ads, ads_title, templatepath, langCode)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if
		
		'******************************************************************************
		'		Predispongo la lista di indirizzi email per l'invio della conferma;
		'		Gli indirizzi si distinguono per zona geografica, in base al campo
		'		"Country" selezionato dell'utente viene scelta una mail o un'altra
		'******************************************************************************
		Dim objListaMail, strMailBcc
		strMailBcc = Application("mail_bcc")		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
		
		With iMsg
			Set .Configuration = iConf
				.To = mailTo
				.From = Application("mail_sender")
				.Sender = Application("mail_sender")
				if(Trim(strMailBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.ads.mail.subject.label.ads_contact")&" "&ads_title
				.CreateMHTMLBody strServerName&Application("baseroot")&Application("dir_upload_templ")&templatepath&"?lang_code="&langCode&"&id_ads="&id_ads&"&mailText="&mailText&"&sid="&Rnd(session.SessionID), 0				
				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub

	Public Sub sendMailVoucher(voucher_code, is_gift, strEmail, templatepath, moduleParams, langCode)
		on error resume next

		Dim langMail		
		Set langMail = new LanguageClass
		'*** verifico se è stata passata la lingua dell'utente e la imposto come langMail.setLangCode(xxx)
		if not(isNull(langCode)) AND not(langCode ="") AND not(langCode="null")  then
			langMail.setLangCode(langCode)
			langMail.setLangElements(langMail.getListaElementsByLang(langMail.getLangCode()))
		end if

		Dim objListaMail, strMailCc, strMailBcc
		strMailBcc = Application("mail_receiver")		
		
		Dim strServerName, objMail, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlgetlatestversion") = True
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		else
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		end if
		
		iConf.Fields.Update
		
		if(isNull(templatepath) OR templatepath="")then
			templatepath=strServerName&Application("baseroot") & "/editor/voucher/mail_notif/mail_voucher_guest.asp"
		else
			templatepath=strServerName&Application("baseroot")&Application("dir_upload_templ")&templatepath
		end if
													
		With iMsg
			Set .Configuration = iConf
				.To = strEmail
				.From = Application("mail_sender")
				.Sender = Application("mail_sender")
				if(Trim(strMailBcc) <> "") then .Bcc = strMailBcc end if
				.Subject = langMail.getTranslated("backend.voucher.mail.subject.label.new_voucher")
				.CreateMHTMLBody templatepath&"?is_gift="&is_gift&"&voucher_code="&voucher_code&"&module_params="&moduleParams&"&lang_code="&langCode&"&sid="&Rnd(session.SessionID), 0

				if not(Application("mail_server") = "") then
					.Send				
				end if
		End With
		
		Set langMail = nothing
	
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub

End Class


'***************************** TEST DI ESEMPIO *****************************
'Dim MiaMail
'Set MiaMail = Server.CreateObject("CDO.Message")
'MiaMail.From = "blackhole01@gmail.com"
'MiaMail.To = "d.testa@sol-tec.it"
'MiaMail.Cc = "altroindirizzo@aruba.it;ancora@aruba.it"
'MiaMail.Bcc = "altroindirizzo@aruba.it;ancora@aruba.it"
'MiaMail.Subject = "Invio tramite cdosys"
'MiaMail.TextBody = "Invia tramite CDOSYS paragone con cdonts "
'MiaMail.AddAttachment "d:\inetpub\webs\tuodominiocom\file.zip"
'MiaMail.Fields("urn:schemas:httpmail:importance").Value = 2
'MiaMail.Fields.Update()
'MiaMail.Send()
'Set MiaMail = Nothing
'***************************** FINE TEST DI ESEMPIO *****************************
%>