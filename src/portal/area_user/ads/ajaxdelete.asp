<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if(strComp(Cint(strRuoloLogged), Application("guest_role"), 1) = 0) then

		Dim objLogger
		Set objLogger = New LogClass

		Dim objtype, id_objref
		objtype = request("objtype")
		id_objref = request("id_objref")

		On Error Resume Next
		Dim objRef, objTmp, objDict
		Select Case objtype
			Case "content"
				Set objRef = New NewsClass
				call objRef.deleteNews(id_objref)				
				'call objLogger.write("content deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case Else			
		End Select
		
		if(Err.number<>0) then
			response.write(err.description)
		end if
		
		Set objLogger = nothing
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>
