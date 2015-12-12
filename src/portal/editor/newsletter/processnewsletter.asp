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
	
	Dim id_newsletter, strDescrizione, iStato, bolDelNewsletter, strTemplate
	id_newsletter = request("id_newsletter")
	strDescrizione = request("descrizione")
	iStato = request("stato")
	strTemplate = request("template")
	bolDelNewsletter = request("delete_newsletter")
	id_voucher_campaign = request("voucher")
	
	Dim objNewsletter
	Set objNewsletter = New NewsletterClass
	
	if (Cint(id_newsletter) <> -1) then
		if(strComp(bolDelNewsletter, "del", 1) = 0) then
			if(objNewsletter.findNewsletterAssociations(id_newsletter)) then
				response.Redirect(Application("baseroot")&Application("error_page")&"?error=028")		
			else
				call objNewsletter.deleteNewsletter(id_newsletter)
				response.Redirect(Application("baseroot")&"/editor/newsletter/ListaNewsletter.asp")	
			end if
		
		end if
		
	
		call objNewsletter.modifyNewsletter(id_newsletter, strDescrizione, iStato, strTemplate, id_voucher_campaign)
		Set objNewsletter = nothing
		response.Redirect(Application("baseroot")&"/editor/newsletter/ListaNewsletter.asp")		
	else
		call objNewsletter.insertNewsletter(strDescrizione, iStato, strTemplate, id_voucher_campaign)
		Set objNewsletter = nothing
		response.Redirect(Application("baseroot")&"/editor/newsletter/ListaNewsletter.asp")				
	end if

	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>