<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<%
'<!--nsys-rep5-->
%>
<!-- #include virtual="/common/include/Objects/UserGroupClass.asp" -->
<%
'<!---nsys-rep5-->
%>
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->

<%
Response.Buffer = TRUE 
Response.ContentType = "text/csv"
Response.AddHeader "Content-Disposition", "attachment; filename=csv_user.csv"

Dim objUtente, objListaUtenti, objTmpUtenti
Dim counter, tmpObjUsr, iIndex, objNewsletter, objUserNewsletter

order_by = 1
rolef = null
publicf = null
activef = null
if(request("rolef")<>"")then
	rolef = request("rolef")
end if
if(request("publicf")<>"")then
	publicf = request("publicf")
end if
if(request("activef")<>"")then
	activef = request("activef")
end if
if(request("order_by")<>"")then
	order_by = request("order_by")
end if

Set objUtente = New UserClass
Set objListaUtenti = objUtente.findUtente(null, rolef, activef, publicf, 0, order_by)
Set objNewsletter = new NewsletterClass
'<!--nsys-rep6-->
'********** RECUPERO LA LISTA DI GRUPPI UTENTE DISPONIBILI
Dim objGroup
Set objGroup = New UserGroupClass
Dim objDispGroup
On Error Resume Next
Set objDispGroup = objGroup.getListaUserGroup()
if(Err.number <> 0) then
end if
Set objGroup = nothing
'<!---nsys-rep6-->
'********** RECUPERO LA LISTA DI FIELD UTENTE DISPONIBILI
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

userFieldcount =1
hasFieldFilterActive = false
Set objDictFilteredFieldActive = Server.CreateObject("Scripting.Dictionary")
if(hasUserFields) then
	for each k in objListUserField
		Set objField = objListUserField(k)
		On Error Resume next
		if(Cint(objField.getTypeField())=8 OR Cint(objField.getTypeField())=1)then
			Set objFilterfieldValue = objUserField.findFieldMatchValueUnique(objField.getID())
			if(objFilterfieldValue.count>0)then
				if(request(objUserField.getFieldPrefix()&objField.getID())<>"")then
					objDictFilteredFieldActive.add objField.getID(), request(objUserField.getFieldPrefix()&objField.getID())
					hasFieldFilterActive = true
				end if
			end if
			Set objFilterfieldValue = nothing
		end if

		if(Err.number<>0) then
		'response.write(Err.description)
		end if
		Set objField = nothing
		userFieldcount=userFieldcount+1
	next
end if

if(hasFieldFilterActive)then
	for each k in objListaUtenti
		doRemove=true
		for each i in objDictFilteredFieldActive
			valuetmp = objUserField.findFieldMatchValue(i, k)
			'response.write("k:"&k&" - valuetmp:"&valuetmp&" - objDictFilteredFieldActive(i):"&objDictFilteredFieldActive(i)&" - i:"&i&" - check:"& (valuetmp<>objDictFilteredFieldActive(i)))
			'if(valuetmp<>objDictFilteredFieldActive(i))then							
			'	objListaUtenti.remove(k)
			'	exit for
			'end if
			arrFilteredField = Split(objDictFilteredFieldActive(i), ",", -1, 1)
			for cf = 0 to Ubound(arrFilteredField)
				'response.write(" - k:"&k&"; -valuetmp:"&valuetmp&"; -arrFilteredField(cf):"&Trim(arrFilteredField(cf)&"; -equals:"& (valuetmp=Trim(arrFilteredField(cf)))))
				if(valuetmp=Trim(arrFilteredField(cf)))then	
					doRemove=false
					exit for
				end if							
			next
		next
		if(doRemove)then							
			objListaUtenti.remove(k)
		end if
	next
end if
iIndex = objListaUtenti.Count
objTmpUtenti = objListaUtenti.Items
'response.write("iIndex:"&iIndex)
'response.end
%>

		<%=UCase(langEditor.getTranslated("backend.utenti.detail.table.label.username"))&","%>
		<%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.email"))&","%> 
		<%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.user_role"))&","%>
		<%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.user_active"))&","%>
<%
'<!--nsys-rep7-->
%>
		<%=UCase(langEditor.getTranslated("backend.utenti.detail.table.label.user_group"))&","%>
<%
'<!---nsys-rep7-->
%>
		<%=UCase(langEditor.getTranslated("backend.utenti.detail.table.label.public_profile"))&","%>
		<%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.sconto"))&","%>
		<%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.admin_comments"))&","%>	
		<%=UCase(langEditor.getTranslated("backend.utenti.detail.table.label.subscribe_newsletter"))&","%>		
		<%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.date_insert"))&","%>	
		<%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.date_modify"))%>

		<%
		if(hasUserFields) then
			On Error Resume next
			for each k in objListUserField
				Set objField = objListUserField(k)%>
				<%=","%><%if not(langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription())="") then response.write(UCase(langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription()))) else response.write(UCase(objField.getDescription())) end if%>
			<%next

			if(Err.number<>0) then
			'response.write(Err.description)
			end if
		end if%>
		<%=","&vbCrLf%>
	
<%for counter = 0 to iIndex-1%>
	<%Set tmpObjUsr = objTmpUtenti(counter)
	Dim strRuolo, strUserActive
	
	if(tmpObjUsr.getRuolo() = Application("admin_role")) then
		strRuolo = Application("admin_const")
	elseif (tmpObjUsr.getRuolo() = Application("editor_role")) then
		strRuolo = Application("editor_const")
	else
		strRuolo = Application("guest_const")
	end if
	
	if(tmpObjUsr.getUserActive() = 1) then
		strUserActive = langEditor.getTranslated("backend.commons.yes")
	else
		strUserActive = langEditor.getTranslated("backend.commons.no")
	end if
	%>

		<%response.write("""")%><%=tmpObjUsr.getUserName()%><%=""","%>
		<%response.write("""")%><%=tmpObjUsr.getEmail()%><%=""","%>
		<%response.write("""")%><%=strRuolo%><%=""","%>
		<%response.write("""")%><%=strUserActive%><%=""","%>
<%
'<!--nsys-rep8-->
%>
		<%response.write("""")%><%
		if (Instr(1, typename(objDispGroup), "dictionary", 1) > 0) then
		for each x in objDispGroup
		if (tmpObjUsr.getGroup() = x) then 
			response.Write(objDispGroup(x).getShortDesc())
			exit for
		end if
		next
		end if%><%=""","%>
<%		
'<!---nsys-rep8-->
%>
		<%response.write("""")%><%if(tmpObjUsr.getPublic() = 1) then
			response.Write(langEditor.getTranslated("backend.commons.yes"))
		else
			response.Write(langEditor.getTranslated("backend.commons.no"))
		end if%><%=""","%>
		<%response.write("""")%><%=tmpObjUsr.getSconto()%><%=""","%>
		<%response.write("""")%><%=tmpObjUsr.getAdminComments()%><%=""","%>
		<%
			On Error Resume Next
			Set objUserNewsletter = objNewsletter.getListNewsletterPerUser(tmpObjUsr.getUserID())
			
			response.write("""")
			
			if(strComp(typename(objUserNewsletter), "Dictionary", 1) = 0) then
				for each y in objUserNewsletter
					response.write(y&"; ")
				next
			end if
			
			response.write(""",")
			
			Set objUserNewsletter = nothing
			
			if(Err.number <> 0) then
				'response.write(Err.description)
			end if
		%>					
		<%response.write("""")%><%=tmpObjUsr.getInsertDate()%><%=""","%>	
		<%response.write("""")%><%=tmpObjUsr.getModifyDate()%><%=""","%>

		<%
		if(hasUserFields) then
			'On Error Resume next
			for each k in objListUserField
				Set objField = objListUserField(k)%>
				<%response.write("""")%>
				<%on error resume next
				Set fieldMatchValue = objUserField.findFieldMatch(objField.getID(),tmpObjUsr.getUserID())
				if (Instr(1, typename(fieldMatchValue), "dictionary", 1) > 0) then
					fieldMatchValue = fieldMatchValue.Item("value")
				else
					fieldMatchValue = ""				
				end if
				response.write(fieldMatchValue)
				if Err.number <> 0 then
					'response.write(Err.description)
				end if%>
				<%=""","%>
			<%next

			'if(Err.number<>0) then
			'response.write(Err.description)
			'end if
		end if
		%>		
		<%=vbCrLf%>

	<%Set tmpObjUsr = nothing
next


Set objListUserField = nothing
Set objUserField = nothing
Set objNewsletter = nothing
Set objListaUtenti = nothing
Set objUtente = Nothing%>