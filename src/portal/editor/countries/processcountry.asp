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
	
	Dim id_country, country_code, state_region_code, country_description, state_region_description, active, use_for
	id_country = request("id_country")
	country_code = request("country_code")
	state_region_code = request("state_region_code")
	country_description = request("country_description")
	state_region_description = request("state_region_description")
	active = request("active")
	use_for = request("use_for")
	bolDelCountry = request("delete_country")
	
	Dim objCountry
	Set objCountry = New CountryClass
	
	if (Cint(id_country) <> -1) then
		if(strComp(bolDelCountry, "del", 1) = 0) then
			call objCountry.deleteCountry(id_country)
			response.Redirect(Application("baseroot")&"/editor/countries/ListaCountry.asp")			
		end if		
	
		call objCountry.modifyCountry(id_country,country_code, country_description, state_region_code, state_region_description, active, use_for)
		Set objCountry = nothing
		response.Redirect(Application("baseroot")&"/editor/countries/ListaCountry.asp")		
	else
		newMaxID = objCountry.insertCountry(country_code, country_description, state_region_code, state_region_description, active, use_for)
		
		'/**
		'* aggiorno le localizzazioni se sono state inserite prima di salvare il contenuto
		'*/
		if(request("pregeoloc_el_id")<>"") then
			Set objLoc = new LocalizationClass
			Set listOfPoints = objLoc.findPointByElement(request("pregeoloc_el_id"), 3)
			for each q in listOfPoints
				call objLoc.modifyPointNoTransaction(q, newMaxID, listOfPoints(q).getLatitude(), listOfPoints(q).getLongitude(), listOfPoints(q).getInfo())
			next
			Set listOfPoints = nothing
			Set objLoc = nothing
		end if
		
		Set objCountry = nothing
		response.Redirect(Application("baseroot")&"/editor/countries/ListaCountry.asp")				
	end if

	Set objUserLogged = nothing
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>