<%
if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
Set objUserLoggedTmp = nothing
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing

'/**
'* recupero i valori della news selezionata se id_newsletter <> -1
'*/
Dim id_newsletter, strDescrizione, iStato, strTemplate, objNewsletter, listTemplate, id_voucher_campaign
id_newsletter = request("id_newsletter")
strDescrizione = ""
iStato = 0
strTemplate = ""
id_voucher_campaign = ""

Set objNewsletter = New NewsletterClass
listTemplate = objNewsletter.getListaTemplateNewsletter()

if (Cint(id_newsletter) <> -1) then
	Dim objSelNewsletter
	Set objSelNewsletter = objNewsletter.findNewsletterByID(id_newsletter)
	Set objNewsletter = nothing
	
	id_newsletter = objSelNewsletter.getNewsletterID()
	strDescrizione = objSelNewsletter.getDescrizione()		
	iStato = objSelNewsletter.getStato()	
	strTemplate = objSelNewsletter.getTemplate()
	id_voucher_campaign =  objSelNewsletter.getVoucher()
	Set objSelNewsletter = Nothing
end if
%>