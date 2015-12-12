<%On Error Resume Next%>
<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->
<%
'/**
'* recupero i valori della news selezionata se id_prod <> -1
'*/
Dim strUsername,strPassword,strCognome,strNome,strEmail,strCodFiscPiva,strCountry
Dim strCity,strAddress,strZipCode,numTelephone,numFax,strCompanyName,strBusinessActivity,strWebsite
Dim id_user, registrationCode, birthday, sex, interests, strListOthers
id_user = request("id_utente")

'*** verifico se � stata passata la lingua dell'utente e la imposto come langEditor.setLangCode(xxx)
if not(isNull(request("lang_code"))) AND not(request("lang_code")="") AND not(request("lang_code")="null")  then
	langEditor.setLangCode(request("lang_code"))
	langEditor.setLangElements(langEditor.getListaElementsByLang(langEditor.getLangCode()))
end if

strUsername=""
strPassword=request("password")
strEmail=""


Dim objUtente, objTmpUser
Set objUtente = New UserClass

if not (isNull(id_user)) then
	Set objTmpUser = objUtente.findUserByID(id_user)	
	strUsername=objTmpUser.getUserName()
	strEmail=objTmpUser.getEmail()
	Set objTmpUser = nothing
else
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")			
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=langEditor.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<link rel="stylesheet" href="<%="http://" & request.ServerVariables("SERVER_NAME") & Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/public/layout/include/header.inc" -->
	<div id="backend-content">
		<span class="labelMailSend"><%=langEditor.getTranslated("backend.utenti.mail.label.intro")%></span><br><br>
		<%=langEditor.getTranslated("backend.utenti.mail.label.intro_detail")%><br><br>
		
		<span class="labelMailSend"><%=langEditor.getTranslated("backend.utenti.mail.label.username")%>:</span>&nbsp;<%=strUsername%><br><br>
		<span class="labelMailSend"><%=langEditor.getTranslated("backend.utenti.mail.label.password")%>:</span>&nbsp;<%=strPassword%><br><br>
		<span class="labelMailSend"><%=langEditor.getTranslated("backend.utenti.mail.label.email")%>:</span>&nbsp;<%=strEmail%><br><br>

		<%
		On Error Resume next
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
		
		if(hasUserFields) then
			for each k in objListUserField
				Set objField = objListUserField(k)%>
				<span class="labelMailSend"><%if not(langEditor.getTranslated("backend.utenti.mail.label."&objField.getDescription())="") then response.write(langEditor.getTranslated("backend.utenti.mail.label."&objField.getDescription())) else response.write(objField.getDescription()) end if%>:</span>&nbsp;
				<%
				On Error Resume next
				Set objFieldValue=objUserField.findFieldMatch(objField.getID(), id_user)
				fieldValue=objFieldValue.Item("value")
				
				if(CInt(objField.getTypeField())=4) then
					label = fieldValue
					if(CInt(objField.getTypeContent())=5) then
						if not(langEditor.getTranslated("portal.commons.select.option.country."&fieldValue)="") then label=langEditor.getTranslated("portal.commons.select.option.country."&fieldValue) end if
					else
						if not(langEditor.getTranslated("portal.commons.user_field.label."&fieldValue)="") then label=langEditor.getTranslated("portal.commons.user_field.label."&fieldValue) end if
					end if
					response.write(label)
				elseif(CInt(objField.getTypeField())=5 OR CInt(objField.getTypeField())=6 OR CInt(objField.getTypeField())=7) then
					label = ""
					fieldValue = split(fieldValue,",")
					for each y in fieldValue
						if not(langEditor.getTranslated("portal.commons.user_field.label."&y)="") then 
							label=label&langEditor.getTranslated("portal.commons.user_field.label."&y)&",&nbsp;"
						else
							label=label&y&",&nbsp;"
						end if
					next
					if(Len(label) > 0) then label = Left(label,(Len(label)-7))
					response.write(label)
				else
					label = fieldValue
					if not(langEditor.getTranslated("portal.commons.user_field.label."&fieldValue)="") then label=langEditor.getTranslated("portal.commons.user_field.label."&fieldValue) end if
					response.write(label)
				end if
				
				response.write("<br><br>")
				
				if(Err.number<>0) then
					'response.write(Err.description)
				end if
			next
		end if

		Set objListUserField = nothing
		Set objUserField = nothing
		%>
	</div>
</div>
</body>
</html>
<%
Set objUtente = nothing

if(Err.number <> 0) then
	'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>
