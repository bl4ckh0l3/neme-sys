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
'* recupero i valori della news selezionata se id_country <> -1
'*/
Dim id_country, country_code, state_region_code, country_description, state_region_description, active, use_for
id_country = request("id_country")
country_code = ""
state_region_code = ""
country_description = ""
state_region_description = ""
active = ""
use_for = ""

if (Cint(id_country) <> -1) then
	Dim objCountry, objSelCountry
	Set objCountry = New CountryClass
	Set objSelCountry = objCountry.findCountryByID(id_country)
	Set objCountry = nothing
	
	id_country = objSelCountry.getID()
	country_code = objSelCountry.getCountryCode()
	state_region_code = objSelCountry.getStateRegionCode()
	country_description = objSelCountry.getCountryDescription()
	state_region_description = objSelCountry.getStateRegionDescription()
	active = objSelCountry.isActive()
	use_for = objSelCountry.getUseFor()
end if
%>