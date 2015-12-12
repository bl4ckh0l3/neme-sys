<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Type Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->
<%
Class NewsletterClass
	Private id_newsletter
	Private id_utente
	Private descrizione
	Private stato
	Private template
	Private voucher
	
	
	Public Function getNewsletterID()
		getNewsletterID = id_newsletter
	End Function
	
	Public Sub setNewsletterID(strID)
		id_newsletter = strID
	End Sub	
	
	Public Function getUserID()
		getUserID = id_utente
	End Function
	
	Public Sub setUserID(strUtente)
		id_utente = strUtente
	End Sub
	
	Public Function getDescrizione()
		getDescrizione = descrizione
	End Function
	
	Public Sub setDescrizione(strDescrizione)
		descrizione = strDescrizione
	End Sub
	
	Public Function getStato()
		getStato = stato
	End Function
	
	Public Sub setStato(strStato)
		stato = strStato
	End Sub
	
	Public Function getTemplate()
		getTemplate = template
	End Function
	
	Public Sub setTemplate(strTemplate)
		template = strTemplate
	End Sub
	
	Public Function getVoucher()
		getVoucher = voucher
	End Function
	
	Public Sub setVoucher(strVoucher)
		voucher = strVoucher
	End Sub
	
	
	Public Function getMaxIDNewsletter()
		on error resume next
		
		getMaxIDNewsletter = -1
		
		Dim objDB, strSQL, objRS, objConn
		strSQL = "SELECT max(newsletter.id_newsletter) as id FROM newsletter;"

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Set objRS = objConn.Execute(strSQL)
		if objRS.EOF then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")		
		else
			getMaxIDNewsletter = objRS("id")	
		end if
				
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
			
	Public Function insertNewsletter(strDescrizione, intStato, strTemplate, id_voucher_campaign)
		on error resume next
		Dim test
		insertNewsletter = -1
		
		Dim objDB, strSQL, objRS, objConn		
		
		strSQL = "INSERT INTO newsletter(descrizione, stato, template, id_voucher_campaign) VALUES("
		strSQL = strSQL & "?,?,?,"
		if(isNull(id_voucher_campaign) OR id_voucher_campaign = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ");"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,intStato)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTemplate)
		if not isNull(id_voucher_campaign) AND not(id_voucher_campaign = "") then
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_voucher_campaign)
		end if
		objCommand.Execute()
		Set objCommand = Nothing

		Set objRS = objConn.Execute("SELECT max(newsletter.id_newsletter) as id FROM newsletter")
		if not (objRS.EOF) then
			insertNewsletter = objRS("id")	
		end if		
		Set objRS = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyNewsletter(id, strDescrizione, intStato, strTemplate, id_voucher_campaign)
		on error resume next
		Dim objDB, strSQL, objRS, objConn
				
		strSQL = "UPDATE newsletter SET "
		strSQL = strSQL & "descrizione=?,"
		strSQL = strSQL & "stato=?,"
		strSQL = strSQL & "template=?,"
		if(isNull(id_voucher_campaign) OR id_voucher_campaign = "") then
			strSQL = strSQL & "id_voucher_campaign=NULL"
		else
			strSQL = strSQL & "id_voucher_campaign=?"
		end if
		strSQL = strSQL & " WHERE id_newsletter=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strDescrizione)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,intStato)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,strTemplate)
		if not isNull(id_voucher_campaign) AND not(id_voucher_campaign = "") then
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_voucher_campaign)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub		
		
	Public Sub deleteNewsletter(id)
		on error resume next
		Dim objDB, strSQLDelUserNews, strSQLDelNewsletter, objRS, objConn
		
		'** 
		'* La cancellazione manuale delle tabelle correlate non viene usata
		'* grazie alle relazioni impostate su Access
		'* impostare le stesse relazioni su un altro DB (MySQL)
		'* o riattivare la cancellazione manuale, ma si perde ACID
		
		strSQLDelUserNews = "DELETE FROM newsletter_x_utente WHERE id_newsletter=?;"
		strSQLDelNewsletter = "DELETE FROM newsletter WHERE id_newsletter=?;" 

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		

		Dim objCommand, objCommand2
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQLDelUserNews
		objCommand2.CommandText = strSQLDelNewsletter
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,19,1,,id)
		objCommand.Execute()
		objCommand2.Execute()	
		Set objCommand = Nothing
		Set objCommand2 = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	
	
	Public Function getListaNewsletter(active)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListaNewsletter = null		
		strSQL = "SELECT * FROM newsletter"
		
		if (isNull(active)) then
			strSQL = "SELECT * FROM newsletter"
		else
			strSQL = strSQL & " WHERE"
			if not(isNull(active)) then strSQL = strSQL & " AND stato=?"
		end if
		
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)
		strSQL = strSQL & " ORDER BY descrizione;" 
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if not(isNull(active)) then objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,active)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objNewsletter
			do while not objRS.EOF				
				Set objNewsletter = new NewsletterClass
				strID = objRS("id_newsletter")
				objNewsletter.setNewsletterID(strID)
				objNewsletter.setDescrizione(objRS("descrizione"))	
				objNewsletter.setStato(objRS("stato"))	
				objNewsletter.setTemplate(objRS("template"))	
				objNewsletter.setVoucher(objRS("id_voucher_campaign"))
				objDict.add strID, objNewsletter
				objRS.moveNext()
			loop
			Set objNewsletter = nothing							
			Set getListaNewsletter = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	Public Function findNewsletterByID(id_newsletter)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		findNewsletterByID = null		
		strSQL = "SELECT * FROM newsletter WHERE id_newsletter=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_newsletter)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then
			Dim objNewsletter
			Set objNewsletter = new NewsletterClass
				objNewsletter.setNewsletterID(objRS("id_newsletter"))
				objNewsletter.setDescrizione(objRS("descrizione"))	
				objNewsletter.setStato(objRS("stato"))	
				objNewsletter.setTemplate(objRS("template"))		
				objNewsletter.setVoucher(objRS("id_voucher_campaign"))					
			Set findNewsletterByID = objNewsletter
			Set objNewsletter = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function getUserPerNewsletter(id_newsletter)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getUserPerNewsletter = null		
		strSQL = "SELECT * FROM newsletter_x_utente WHERE id_newsletter=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_newsletter)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id_newsletter")
				strIDUsr = objRS("id_utente")		
				objDict.add strIDUsr, strID
				objRS.moveNext()
			loop
							
			Set getUserPerNewsletter = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function getListNewsletterPerUser(id_user)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict
		getListNewsletterPerUser = null		
		strSQL = "SELECT n.id_newsletter, n.descrizione FROM newsletter n INNER JOIN newsletter_x_utente ON (newsletter_x_utente.id_newsletter = n.id_newsletter) WHERE id_utente=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_user)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF
				strID = objRS("id_newsletter")
				strIDUsr = objRS("descrizione")	

				objDict.add strIDUsr, strID
				objRS.moveNext()
			loop
							
			Set getListNewsletterPerUser = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findNewsletterAssociations(id_newsletter)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		Dim strSQL2, strSQL3, strSQL4
		findNewsletterAssociations = false	
		strSQL = "SELECT newsletter_x_utente.id_newsletter FROM newsletter_x_utente WHERE newsletter_x_utente.id_newsletter=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_newsletter)
		Set objRS = objCommand.Execute()		
		
		if not(objRS.EOF) then							
			findNewsletterAssociations = true				
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function getListaTemplateNewsletter()
		on error resume next
		getListaTemplateNewsletter = null		

		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		uploadsDirVar = Application("baseroot")&Application("dir_upload_templ")	
		uploadsDirVar = Server.MapPath(uploadsDirVar)	
		uploadsDirVarAsp = uploadsDirVar & "\newsletter\"

		set fs = server.createobject("Scripting.FileSystemObject")
		if (objFSO.FolderExists(uploadsDirVarAsp)) then
			set f = fs.getfolder(uploadsDirVarAsp)
			set fl = f.files    ' list of files
			counter = 1
			
			listTemplate=""
			for each strfile in fl
				listTemplate=listTemplate&strfile.name
				if(counter<Cint(fl.count))then listTemplate=listTemplate&";"
				counter = counter+1
			next
			
			set fl = nothing
			set f = nothing
		end if
		Set objFSO = nothing

		listTemplate = Split(listTemplate, ";", -1, 1)
		if(isArray(listTemplate)) then
			getListaTemplateNewsletter = listTemplate
		end if
 
		if Err.number <> 0 then
			getListaTemplateNewsletter = Split("", " ", -1, 1)	
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Sub sendNewsletter(id_news, id_newsletter)
		on error resume next
		
		Dim strServerName, objMail, objUtente, strListaUtentiNewsletter, HTML, nomeDominio, indirizzoIp, modulo, browserSistemaOperativo
		strServerName = "http://" & request.ServerVariables("SERVER_NAME")

		nomeDominio 				= Request.ServerVariables("HTTP_HOST")
		indirizzoIp					= Request.ServerVariables("REMOTE_ADDR") 
		modulo						= Request.ServerVariables("HTTP_REFERER")
		browserSistemaOperativo		= Request.ServerVariables("HTTP_USER_AGENT")

		'* creo gli oggetti cdosys sul server e li gestisco		
		DIM iMsg, Flds, iConf, IBodyParts
		
		Set iMsg = CreateObject("CDO.Message")
		Set iConf = CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("mail_server")
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		
		if(not(Application("mail_server_usr") = "") AND not(Application("mail_server_pwd") = "")) then
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("mail_server_usr")
			iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("mail_server_pwd")
		end if
		
		iConf.Fields.Update
		
		With iMsg
			Set .Configuration = iConf
				.To = Application("mail_sender")
				.From = Application("mail_sender")
				.Sender = Application("mail_sender")

				'********************************************************
				' gestire il recupero della lista utenti della newsletter
				Set objUtente = new UserClass
				
				strListaUtentiNewsletter = objUtente.getListaUtentiNewsletter(id_newsletter)
				
				if not(strListaUtentiNewsletter = "") then
					.Bcc = strListaUtentiNewsletter
				end if
				Set objUtente = nothing
				'********************************************************
				
				Dim strsubject, tmp
				Set tmp = new NewsletterClass
				.Subject = tmp.findNewsletterByID(id_newsletter).getDescrizione()
				.CreateMHTMLBody strServerName&Application("baseroot") & Application("dir_upload_templ") &"newsletter/"&tmp.findNewsletterByID(id_newsletter).getTemplate()&"?id_news="&id_news&"&sid="&Rnd(session.SessionID), 31
				Set tmp = nothing
				.Send
		End With
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if	
	End Sub

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
End Class
%>