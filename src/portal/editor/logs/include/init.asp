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
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing

Dim objLog, objListaLog
Set objLog = New LogClass


paramDelete = request("delete_log")
if(paramDelete = "1") then		
	call objLog.deleteLogs(request("log_type"),request("dta_from"),request("dta_to"))	
end if

Dim itemsXpage, numPage

if not(request("items") = "") then
	session("logsItems") = request("items")
	itemsXpage = session("logsItems")
	session("logsPage") = 1
else
	if not(session("logsItems") = "") then
		itemsXpage = session("logsItems")
	else
		session("logsItems") = 20
		itemsXpage = session("logsItems")
	end if
end if

if not(request("page") = "") then
	session("logsPage") = request("page")
	numPage = session("logsPage")
else
	if not(session("logsPage") = "") then
		numPage = session("logsPage")
	else
		session("logsPage") = 1
		numPage = session("logsPage")
	end if
end if	

Dim paramType, paramDateFrom, paramDateTo, paramDelete		

if not(request("log_type") = "") then
	session("log_type") = request("log_type")
	paramType = session("log_type")
	session("logsPage") = 1
else
	if not(session("log_type") = "") then
		paramType = session("log_type")
	else
		session("log_type") = ""
		paramType = session("log_type")
	end if
end if
if not(request("dta_from") = "") then
	session("dta_from") = request("dta_from")
	paramDateFrom = session("dta_from")
	session("logsPage") = 1
else
	if not(session("dta_from") = "") then
		paramDateFrom = session("dta_from")
	else
		session("dta_from") = Date()
		paramDateFrom = session("dta_from")
	end if
end if
if not(request("dta_to") = "") then
	session("dta_to") = request("dta_to")
	paramDateTo = session("dta_to")
	session("logsPage") = 1
else
	if not(session("dta_to") = "") then
		paramDateTo = session("dta_to")
	else
		session("dta_to") = Date()
		paramDateTo = session("dta_to")
	end if
end if

if(not(isNull(request("resetMenu"))) AND request("resetMenu") = "1") then
	session("logsPage") = 1
	numPage = session("logsPage")
	session("log_type") = ""
	paramType = session("log_type")
	session("dta_from") = Date()
	paramDateFrom = session("dta_from")
	session("dta_to") = Date()
	paramDateTo = session("dta_to")
end if
%>