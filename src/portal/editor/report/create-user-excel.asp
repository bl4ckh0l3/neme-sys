<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<%
'<!--nsys-rep1-->
%>
<!-- #include virtual="/common/include/Objects/UserGroupClass.asp" -->
<%
'<!---nsys-rep1-->
%>
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->

<%
Response.Buffer = TRUE 
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "filename=excel_user.xls"

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
'<!--nsys-rep2-->
'********** RECUPERO LA LISTA DI GRUPPI UTENTE DISPONIBILI
Dim objGroup
Set objGroup = New UserGroupClass
Dim objDispGroup
On Error Resume Next
Set objDispGroup = objGroup.getListaUserGroup()
if(Err.number <> 0) then
'response.write(Err.description)
end if
Set objGroup = nothing
'<!---nsys-rep2-->
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
%>

<html>
<head>
<title></title>
<style type="text/css"> 
body {
	background: #FFFFFF;
}
.tdHeaderExcel {
	background-color: #432D30;
	text-align: left;
	color: #FFFFFF;
}
</style>
</head>
<body>
<TABLE BORDER=1>
	<tr>
		<td colspan="21"><strong>Totale Utenti:</strong> <%=iIndex%></td>
	</tr>
	<tr>
		<td colspan="21">&nbsp;</td>
	</tr>
	<tr class="tdHeaderExcel">
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.detail.table.label.username"))%></strong></td>
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.email"))%></strong></td> 
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.user_role"))%></strong></td>
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.user_active"))%></strong></td>
<%
'<!--nsys-rep3-->
%>
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.detail.table.label.user_group"))%></strong></td>
<%
'<!---nsys-rep3-->
%>
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.detail.table.label.public_profile"))%></strong></td>
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.sconto"))%></strong></td>
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.admin_comments"))%></strong></td>	
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.detail.table.label.subscribe_newsletter"))%></strong></td>		
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.date_insert"))%></strong></td>	
		<td><strong><%=UCase(langEditor.getTranslated("backend.utenti.include.table.header.date_modify"))%></strong></td>

		<%
		if(hasUserFields) then
			On Error Resume next
			for each k in objListUserField
				Set objField = objListUserField(k)%>
				<td><strong><%if not(langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription())="") then response.write(UCase(langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription()))) else response.write(UCase(objField.getDescription())) end if%></strong></td>
			<%next

			if(Err.number<>0) then
			'response.write(Err.description)
			end if
		end if%>	
	</tr>
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
	<tr>
		<td><%=tmpObjUsr.getUserName()%></td>
		<td><%=tmpObjUsr.getEmail()%></td> 
		<td><%=strRuolo%></td> 
		<td><%=strUserActive%></td>
<%
'<!--nsys-rep4-->
%>
		<td><%
		if (Instr(1, typename(objDispGroup), "dictionary", 1) > 0) then
		for each x in objDispGroup
		if (tmpObjUsr.getGroup() = x) then 
			response.Write(objDispGroup(x).getShortDesc())
			exit for
		end if
		next
		end if%></td>
<%
'<!---nsys-rep4-->
%>
		<td><%if(tmpObjUsr.getPublic() = 1) then
			response.Write(langEditor.getTranslated("backend.commons.yes"))
		else
			response.Write(langEditor.getTranslated("backend.commons.no"))
		end if%></td>
		<td><%=tmpObjUsr.getSconto()%> %</td>
		<td><%=tmpObjUsr.getAdminComments()%></td>	
		<td>
		<%
			On Error Resume Next
			Set objUserNewsletter = objNewsletter.getListNewsletterPerUser(tmpObjUsr.getUserID())
			
			if(strComp(typename(objUserNewsletter), "Dictionary", 1) = 0) then
				for each y in objUserNewsletter
					response.write(y&"; ")
				next
			end if
			
			Set objUserNewsletter = nothing
			
			if(Err.number <> 0) then
				'response.write(Err.description)
			end if
		%>
		</td>			
		<td><%=tmpObjUsr.getInsertDate()%></td>	
		<td><%=tmpObjUsr.getModifyDate()%></td>	

		<%
		if(hasUserFields) then
			'On Error Resume next
			Dim userFieldcount, fieldCssClass
			for each k in objListUserField
				Set objField = objListUserField(k)%>
				<td>
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
				&nbsp;</td>
			<%next

			'if(Err.number<>0) then
			'response.write(Err.description)
			'end if
		end if
		%>		
	</tr>
	<%Set tmpObjUsr = nothing
next%>

</TABLE>
</body>
</html>
<%

Set objListUserField = nothing
Set objUserField = nothing
Set objNewsletter = nothing
Set objListaUtenti = nothing
Set objUtente = Nothing%>