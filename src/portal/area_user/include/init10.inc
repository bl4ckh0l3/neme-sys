<%
if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim id_utente, strUserName, strPwd, strNome, strCognome, strEmail, strUsrRuolo, bolPrivacy, bolNewsletter, objUsrTarget
Dim numTelephone, numFax, strCompanyName, strAddress, strCity, strCountry
Dim strZipCode, strWebsite, strBusinessActivity
Dim bolUserActive, numSconto, strAdminComments, strCodFiscPiva
Dim dateInsertDate, dateModifyDate, objNewsletter, objNewsletterUsr
Dim toDo, usrHasImg
Dim birthday, sex, interests, strListOthers, bolPublic
Dim default_date
default_date = "dd/mm/yyyy"
toDo = request("toDo")


Dim order_photo_by, reqPhotoBy
order_photo_by = 6
reqPhotoBy = request("order_by")

if (not(isNull(reqPhotoBy)) AND not(reqPhotoBy = "")) then
	order_photo_by = reqPhotoBy	
end if

Dim totPages, photoXpage, numPage

photoXpage = 10
numPage = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

numMaxImg = 1
if(not(request("numMaxImgs") = "")) then
	numMaxImg = request("numMaxImgs")
end if


id_utente = -1
strUserName = lang.getTranslated("frontend.area_user.manage.label.username") 
strPwd = ""
strNome = lang.getTranslated("frontend.area_user.manage.label.name")
strCognome = lang.getTranslated("frontend.area_user.manage.label.surname")
strEmail = lang.getTranslated("frontend.area_user.manage.label.email")
strUsrRuolo = Application("guest_role")
bolPrivacy = true
bolNewsletter = true 
numTelephone = lang.getTranslated("frontend.area_user.manage.label.telephone")
numFax = lang.getTranslated("frontend.area_user.manage.label.fax")
strCompanyName = lang.getTranslated("frontend.area_user.manage.label.company_name")
strAddress = lang.getTranslated("frontend.area_user.manage.label.street")
strCity = lang.getTranslated("frontend.area_user.manage.label.city")
strCountry = ""
strZipCode = lang.getTranslated("frontend.area_user.manage.label.zip_code")
strWebsite = lang.getTranslated("frontend.area_user.manage.label.website")
strBusinessActivity = ""
if(Application("confirm_registration") = 1) then
bolUserActive = "1"
else
bolUserActive = "0"
end if
numSconto = 0
strAdminComments = ""
strCodFiscPiva = lang.getTranslated("frontend.area_user.manage.label.cod_fisc_piva")
objUsrTarget = null
dateInsertDate = ""
dateModifyDate = ""
objNewsletterUsr = null
usrHasImg = false
birthday = default_date
sex = ""
interests =  ""
strListOthers = ""
bolPublic = "0"

Set objNewsletter = new NewsletterClass

if not(Session("objUtenteLogged")) then
	Dim objUserLogged, objUserLoggedTmp, objListaRuoli
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objListaRuoli = objUserLoggedTmp.getListaRuoli()
	Set objUserLoggedTmp = nothing
		
	id_utente = objUserLogged.getUserID()
	strUserName = objUserLogged.getUserName()
	strPwd = objUserLogged.getPassword()
	strNome = objUserLogged.getNome()
	if(strNome="")then strNome = lang.getTranslated("frontend.area_user.manage.label.name") end if
	strCognome = objUserLogged.getCognome()
	if(strCognome="")then strCognome = lang.getTranslated("frontend.area_user.manage.label.surname") end if
	strEmail = objUserLogged.getEmail()
	strUsrRuolo = objUserLogged.getRuolo()
	bolPrivacy = objUserLogged.getPrivacy()
	bolNewsletter = objUserLogged.getNewsletter() 
	numTelephone = objUserLogged.getTelephone()
	if(numTelephone="")then numTelephone = lang.getTranslated("frontend.area_user.manage.label.telephone") end if
	numFax = objUserLogged.getFax()
	if(numFax="")then numFax = lang.getTranslated("frontend.area_user.manage.label.fax") end if
	strCompanyName = objUserLogged.getCompanyName()
	if(strCompanyName="")then strCompanyName = lang.getTranslated("frontend.area_user.manage.label.company_name") end if
	strAddress = objUserLogged.getAddress()
	if(strAddress="")then strAddress = lang.getTranslated("frontend.area_user.manage.label.address") end if
	strCity = objUserLogged.getCity()
	if(strCity="")then strCity = lang.getTranslated("frontend.area_user.manage.label.city") end if
	strCountry = objUserLogged.getCountry()
	strZipCode = objUserLogged.getZipCode()
	if(strZipCode="")then strZipCode = lang.getTranslated("frontend.area_user.manage.label.zip_code") end if
	strWebsite = objUserLogged.getWebsite()
	if(strWebsite="")then strWebsite = lang.getTranslated("frontend.area_user.manage.label.website") end if
	strBusinessActivity = objUserLogged.getBusinessActivity()
	bolUserActive = objUserLogged.getUserActive()
	numSconto = objUserLogged.getSconto()
	strAdminComments = objUserLogged.getAdminComments()
	strCodFiscPiva = objUserLogged.getCodFiscPiva()
	if(strCodFiscPiva="")then strCodFiscPiva = lang.getTranslated("frontend.area_user.manage.label.cod_fisc_piva") end if
	dateInsertDate = objUserLogged.getInsertDate()
	dateModifyDate = objUserLogged.getModifyDate()
	usrHasImg = objUserLogged.hasImageUser(id_utente)
	birthday = objUserLogged.getBirthday()
	sex = objUserLogged.getSex()
	interests = objUserLogged.getInterests()
	strListOthers = objUserLogged.getListOthers()
	bolPublic = objUserLogged.getPublic()

	'/*
	' * se l'utente loggato è admin lo mando all'inserimento utenti del content manager
	' */
	if(strComp(Cint(strUsrRuolo), Application("admin_role"), 1) = 0) then
		response.Redirect(Application("baseroot") & "/editor/InserisciUtente.asp?id_utente="& id_utente)
	end if
	
	if not(isNull(objUserLogged.getTargetPerUser(id_utente))) then
		Set objUsrTarget = objUserLogged.getTargetPerUser(id_utente)
	end if		

	if not(isNull(objUserLogged.getNewsletterPerUser(id_utente))) then
		Set objNewsletterUsr = objUserLogged.getNewsletterPerUser(id_utente)
	end if
end if


'********** CREAZIONE ARRAY CON LE VOCI DEI CAMPI SELECT E CHECKBOX
Dim selectBusiness, checkboxInterests, checkboxOther

Set selectBusiness = Server.CreateObject("Scripting.Dictionary")
selectBusiness.add "ba_arte", lang.getTranslated("portal.commons.select.option.activity.ba_arte")
selectBusiness.add "ba_comunicazione", lang.getTranslated("portal.commons.select.option.activity.ba_comunicazione")
selectBusiness.add "ba_costruzioni", lang.getTranslated("portal.commons.select.option.activity.ba_costruzioni")
selectBusiness.add "ba_finanza", lang.getTranslated("portal.commons.select.option.activity.ba_finanza")
selectBusiness.add "ba_informatica", lang.getTranslated("portal.commons.select.option.activity.ba_informatica")
selectBusiness.add "ba_servizi", lang.getTranslated("portal.commons.select.option.activity.ba_servizi")
selectBusiness.add "ba_insurance", lang.getTranslated("portal.commons.select.option.activity.ba_insurance")
selectBusiness.add "ba_others", lang.getTranslated("portal.commons.select.option.activity.ba_others")

Set checkboxInterests = Server.CreateObject("Scripting.Dictionary") 
checkboxInterests.add "in_read", lang.getTranslated("portal.commons.checkbox.interests.in_read")
checkboxInterests.add "in_music", lang.getTranslated("portal.commons.checkbox.interests.in_music")
checkboxInterests.add "in_computer", lang.getTranslated("portal.commons.checkbox.interests.in_computer")
checkboxInterests.add "in_video", lang.getTranslated("portal.commons.checkbox.interests.in_video")
checkboxInterests.add "in_auto", lang.getTranslated("portal.commons.checkbox.interests.in_auto")
checkboxInterests.add "in_viaggi", lang.getTranslated("portal.commons.checkbox.interests.in_viaggi")
checkboxInterests.add "in_moda", lang.getTranslated("portal.commons.checkbox.interests.in_moda")
checkboxInterests.add "in_fotografia", lang.getTranslated("portal.commons.checkbox.interests.in_fotografia")

Set checkboxOther = Server.CreateObject("Scripting.Dictionary")   	
checkboxOther.add "ot_knowby_friend", lang.getTranslated("portal.commons.checkbox.others.ot_knowby_friend")
checkboxOther.add "ot_knowby_search", lang.getTranslated("portal.commons.checkbox.others.ot_knowby_search")
checkboxOther.add "ot_knowby_site", lang.getTranslated("portal.commons.checkbox.others.ot_knowby_site")
checkboxOther.add "ot_knowby_other", lang.getTranslated("portal.commons.checkbox.others.ot_knowby_other")

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>