<!-- #include virtual="/common/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/SendMailClass.asp" -->
<%
Dim id_element, element_type, objCommento, objSelectedCommento, del_commento, posted, id_commento, message, comment_type, active, unlock_commento
id_element = request("id_element")
element_type = request("element_type")
id_commento = request("id_commento")
del_all_commenti = request("del_all_commenti")
del_commento = request("del_commento")
message = request("message")
comment_type = request("comment_type")
active = request("active")
unlock_commento = request("unlock_commento")
posted = false

Set objCommento = New CommentsClass

if(del_all_commenti = 1) then
	call objCommento.deleteCommentiByIDElement(id_element,element_type)
end if

if(del_commento = 1) then
	call objCommento.deleteCommento(id_commento)
end if

if(unlock_commento = 1) then
	call objCommento.updateStatus(id_commento,1)
end if

Dim objUserLogged, objUserLoggedTmp, isLogged, strRuoloLogged, numIdUser
isLogged = false
numIdUser = -1

Set objUserLogged = new UserClass
if not(isEmpty(Session("objUtenteLogged"))) AND not(Session("objUtenteLogged") = "") then
	Set objUserLoggedTmp = objUserLogged.findUserByID(Session("objUtenteLogged"))
	if not(isNull(objUserLoggedTmp)) AND (Instr(1, typename(objUserLoggedTmp), "UserClass", 1) > 0) then 
		strRuoloLogged = objUserLoggedTmp.getRuolo()  
		numIdUser = objUserLoggedTmp.getUserID()
	end if
	Set objUserLoggedTmp = nothing
	isLogged = true
elseif not(isEmpty(Session("objCMSUtenteLogged"))) AND not(Session("objCMSUtenteLogged") = "") then
	Set objCMSUtenteLoggedTmp = objUserLogged.findUserByID(Session("objCMSUtenteLogged"))
	if not(isNull(objCMSUtenteLoggedTmp)) AND (Instr(1, typename(objCMSUtenteLoggedTmp), "UserClass", 1) > 0) then 
		strRuoloLogged = objCMSUtenteLoggedTmp.getRuolo()
		numIdUser = objCMSUtenteLoggedTmp.getUserID() 
	end if
	Set objCMSUtenteLoggedTmp = nothing
	isLogged = true
end if
Set objUserLogged = nothing

Response.Charset="UTF-8"
Session.CodePage  = 65001
		
if(message <> "") then
	message = Replace(message, "'", "&#39;", 1, -1, 1)
	'** sostituisco dal message:
		'èéàòùì
	'** con:
		'&egrave;&eacute;&agrave;&ograve;&ugrave;&igrave;
	message = Replace(message, "è", "&egrave;", 1, -1, 1)
	message = Replace(message, "é", "&eacute;", 1, -1, 1)
	message = Replace(message, "à", "&agrave;", 1, -1, 1)
	message = Replace(message, "ò", "&ograve;", 1, -1, 1)
	message = Replace(message, "ù", "&ugrave;", 1, -1, 1)
	message = Replace(message, "ì", "&igrave;", 1, -1, 1)
		
	Dim newCommentMax

	if (isLogged) then
		newCommentMax =  objCommento.insertCommentoNoTransaction(id_element, element_type, numIdUser, message, comment_type, active)
	else
		newCommentMax =  objCommento.insertCommentoNoTransaction(id_element, element_type, numIdUser, message, comment_type, active)
	end if	
	posted = true
		
	if(Application("use_comments_filter")=1 AND Application("mail_comment_receiver") <> "") then
		'Spedisco la mail di conferma registrazione
		Dim objMail
		Set objMail = New SendMailClass
		call objMail.sendMailComment(newCommentMax, Application("mail_comment_receiver"), Application("str_editor_lang_code_default"))
		Set objMail = Nothing
	end if
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<script language="JavaScript">
function sendForm(){
	if(document.send_commento.message.value == ""){
		alert("<%=lang.getTranslated("frontend.popup.js.alert.insert_commento")%>");
		return;
	}else{
		document.send_commento.submit();	
	}
}

function refreshParentData() {
	opener.document.forms["form_reload_page"].submit();
}
</script>
</head>
<body onunload="refreshParentData();" <%if(posted AND element_type=2)then response.write("onload=window.close();") end if%>>
	<div id="container">
		<div id="content-popup">	
		<%
		if(not(isNull(objCommento.findCommentiByIDElement(id_element, element_type, null)))) AND (Instr(1, typename(objCommento.findCommentiByIDElement(id_element, element_type, null)), "dictionary", 1) > 0) then
			Dim  x, objTmpCommento		
			if (isLogged) then
				if (CInt(strRuoloLogged) = Application("admin_role")) then
					Set objSelectedCommento = objCommento.findCommentiByIDElement(id_element, element_type, null)
					
					for each x in objSelectedCommento.Keys
						Set objTmpCommento = objSelectedCommento(x)
						response.write(objTmpCommento.getDtaInserimento())		
						
						if(objTmpCommento.getActive()=0) then%>
						<a href="<%=Application("baseroot") & "/public/layout/include/popupInsertComments.asp?unlock_commento=1&id_commento="&objTmpCommento.getIDCommento()&"&element_type="&element_type&"&id_element=" & id_element%>" title="<%=lang.getTranslated("portal.templates.commons.label.comment_unlock")%>"><img src="<%=Application("baseroot")&"/editor/img/lock_open.png"%>" vspace="0" hspace="2" border="0" align="absmiddle" alt="<%=lang.getTranslated("portal.templates.commons.label.comment_unlock")%>"></a>&nbsp;&nbsp;
						<%end if%>
						<a href="<%=Application("baseroot") & "/public/layout/include/popupInsertComments.asp?del_commento=1&id_commento="&objTmpCommento.getIDCommento()&"&element_type="&element_type&"&id_element=" & id_element%>"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" vspace="0" hspace="2" border="0" align="absmiddle"></a>
						
						<%response.write("<br>" & objTmpCommento.getMessage()&"<br><hr>")
					next
					
					Set objSelectedCommento = nothing
				end if
			end if
		else
			response.Write("<div align='center'>"&lang.getTranslated("frontend.popup.label.no_comments")&"</div><br>")
		end if%>
		<br><br>
		<%if (isLogged) then%>	
			<form name="send_commento" method="post" action="<%=Application("baseroot") & "/public/layout/include/popupInsertComments.asp"%>" accept-charset="UTF-8">
			<input type="hidden" name="id_element" value="<%=id_element%>">
			<input type="hidden" name="element_type" value="<%=element_type%>">
			
			<div><span class="labelForm"><%=lang.getTranslated("frontend.popup.label.insert_commento")%></span><br>
			<textarea class="formFieldTXTTextareaComment" name="message" id="message" onclick="$('#message').focus();"></textarea></div>
			
			<div><span><%=lang.getTranslated("frontend.area_user.manage.label.like")%></span><br>
			<select name="comment_type" id="comment_type">
				<OPTION VALUE="1"><%=lang.getTranslated("portal.commons.yes")%></OPTION>
				<OPTION VALUE="0"><%=lang.getTranslated("portal.commons.no")%></OPTION>
			</select>			
			</div>
			
			<%if(Application("use_comments_filter")=1) then
				if(CInt(strRuoloLogged) = Application("admin_role")) then%>
				<div><span><%=lang.getTranslated("frontend.area_user.manage.label.active")%></span><br>
				<select name="active" id="active">
					<OPTION VALUE="1"><%=lang.getTranslated("portal.commons.yes")%></OPTION>
					<OPTION VALUE="0"><%=lang.getTranslated("portal.commons.no")%></OPTION>
				</select>			
				</div>
				<%else%>
				<input type="hidden" name="active" value="0">
				<%end if
			else%>
				<input type="hidden" name="active" value="1">			
			<%end if%>
			
			<br><input type="button" name="send" value="<%=lang.getTranslated("frontend.popup.label.insert_commento")%>" onclick="javascript:sendForm();"><!--<a href="javascript:sendForm();"><img src=<%'=Application("baseroot")&"/common/img/add.png"%> vspace="2" hspace="2" border="0" align="middle" alt="<%'=lang.getTranslated("frontend.popup.alt.label.add_comments")%>"><%'=lang.getTranslated("frontend.popup.label.insert_commento")%></a>-->
			<%if(CInt(strRuoloLogged) = Application("admin_role")) then%>
				&nbsp;&nbsp;<input type="button" name="delete" value="<%=lang.getTranslated("frontend.popup.alt.label.delete_comments")%>" onclick="location.href='<%=Application("baseroot") & "/common/include/popupInsertComments.asp?del_all_commenti=1&id_element="&id_element&"&element_type="&element_type%>';"><!--<a href="<%=Application("baseroot") & "/common/include/popupInsertComments.asp?del_all_commenti=1&id_element="&id_element&"&element_type="&element_type%>"><img src=<%'=Application("baseroot")&"/common/img/cancel.png"%> vspace="2" hspace="2" border="0" align="middle" alt="<%'=lang.getTranslated("frontend.popup.alt.label.delete_comments")%>"><%'=lang.getTranslated("frontend.popup.alt.label.delete_comments")%></a>-->
			<%end if%>
			</form>
		<%else
			response.Write(lang.getTranslated("frontend.login.label.login_needed")&"<br>")
		end if
		
		if(posted) then
			response.write(lang.getTranslated("frontend.popup.label.comment_posted"))
		end if%>
		<div align="center" style="padding-top:30px;"><a href="javascript:refreshParentData();window.close();"><%=lang.getTranslated("frontend.popup.label.close_window")%></a></div><br>
		</div>
	</div>
</body>
</html>
<%Set objCommento = nothing%>