<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/SendMailClass.asp" -->
<!-- include virtual="/common/include/captcha/adovbs.asp"-->
<!-- #include virtual="/common/include/captcha/iasutil.asp"-->
<!-- #include virtual="/common/include/captcha/functions.asp"-->
<%
function TestCaptcha(byval valSession, byval valCaptcha)
	dim tmpSession
	valSession = Trim(valSession)
	valCaptcha = Trim(valCaptcha)
	if (valSession = vbNullString) or (valCaptcha = vbNullString) then
		TestCaptcha = false
	else
		tmpSession = valSession
		valSession = Trim(Session(valSession))
		Session(tmpSession) = vbNullString

		if valSession = vbNullString then
			TestCaptcha = false
		else
			valCaptcha = Replace(valCaptcha,"i","I")
			if StrComp(valSession,valCaptcha,1) = 0 then
				TestCaptcha = true
			else
				TestCaptcha = false
			end if
		end if		
	end if
end function
	
if(Cint(request("mail_sent")) = 1) then	
	if(Application("use_recaptcha") = 0) then
		'************* FUNZIONE PER IL VECCHIO CAPTCHA
		' verifico che il codice captcha inserito dall'utente corrisponda con il captcha generato
		' in caso contrario rimando alla pagine di registrazione con errore
		' devo usare campo hidden perch� originale ( request("captchacode") ) non viene recuperato
		if not TestCaptcha("ASPCAPTCHA",  request("sent_captchacode")) then
			response.Redirect(Application("baseroot")&"/public/layout/include/tellafriend.asp?captcha_err=1")
		end if
	else
		'************* RECUPERO PARAMETRI RECAPTCHA
		Dim recaptcha_challenge_field, recaptcha_response_field, recaptcha_private_key, recaptcha_public_key, cTemp
		recaptcha_challenge_field  = request("sent_recaptcha_challenge_field")
		recaptcha_response_field   =request("sent_recaptcha_response_field")
		recaptcha_private_key      = Application("recaptcha_priv_key")
		recaptcha_public_key       = Application("recaptcha_pub_key")

		'************* CHECK VALORE RECAPTCHA
		cTemp = recaptcha_confirm(recaptcha_private_key, recaptcha_challenge_field, recaptcha_response_field)
		If cTemp <> "" Then 
			response.Redirect(Application("baseroot")&"/public/layout/include/tellafriend.asp?captcha_err=1")
		end if
	end if


	Dim objMail, userMail, tellmail, pageURL, tellafriendMessage
	userMail = request("userMail")
	tellmail = request("tellmail")
	pageURL = request("pageURL")
	tellafriendMessage = request("tellafriendMessage")
	
	' invio mail
	Set objMail = New SendMailClass	
	call objMail.sendMailTellaFriend(userMail, tellmail, pageURL, tellafriendMessage, lang.getLangCode())
	Set objMail = Nothing
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<script language="JavaScript">
<!--

function controllaCampiInput(){
	if(document.form_inserisci.userMail.value == ""){
		alert("<%=lang.getTranslated("frontend.tellafriend.js.alert.insert_user_mail")%>");
		document.form_inserisci.userMail.focus();
		return false;
	}	
	if(document.form_inserisci.tellmail.value == ""){
		alert("<%=lang.getTranslated("frontend.tellafriend.js.alert.insert_mail_to")%>");
		document.form_inserisci.tellmail.focus();
		return false;
	}	

	<%if(Application("use_recaptcha") = 0) then%>
		if(document.form_inserisci.captchacode.value == ""){
			alert("<%=lang.getTranslated("frontend.tellafriend.js.alert.insert_captchacode")%>");
			document.form_inserisci.captchacode.focus();
			return false;
		}
		
		// imposto campo hidden sent_captchacode 
		// perch� quello originale non viene recuperato in process
		document.form_inserisci.sent_captchacode.value = document.form_inserisci.captchacode.value;
	<%else%>
		// FUNZIONE PER RECAPTCHA  
		if(document.form_inserisci.recaptcha_response_field.value == ""){
		alert("<%=lang.getTranslated("frontend.tellafriend.js.alert.insert_captchacode")%>");
		document.form_inserisci.recaptcha_response_field.focus();
		return false;
		}
		// imposto campo hidden sent_recaptcha_challenge_field e  sent_recaptcha_response_field
		// perch� quello originale non viene recuperato in process
		document.form_inserisci.sent_recaptcha_challenge_field.value = document.form_inserisci.recaptcha_challenge_field.value;
		document.form_inserisci.sent_recaptcha_response_field.value = document.form_inserisci.recaptcha_response_field.value;
	<%end if%>
  
	document.form_inserisci.submit();
}

function RefreshImage(valImageId) {
	var objImage = document.images[valImageId];
	if (objImage == undefined) {
		return;
	}
	var now = new Date();
	objImage.src = objImage.src.split('?')[0] + '?x=' + now.toUTCString();
}
//-->
</script>
</head>
<body>
	<%
	if(Cint(request("mail_sent")) = 1) then			
		response.write("<div align=center><br/><br/>"&lang.getTranslated("frontend.tellafriend.label.mail_sent")&"</div>")		
	else%>
		<form method="post" name="form_inserisci" action="<%=Application("baseroot")&"/public/layout/include/tellafriend.asp"%>" onsubmit="return controllaCampiInput();" accept-charset="UTF-8">
			<input type="hidden" name="sent_captchacode" value="">
		      <input type="hidden" name="sent_recaptcha_challenge_field" value="">
		      <input type="hidden" name="sent_recaptcha_response_field" value="">
			<input type="hidden" name="mail_sent" value="1" />
			<input type="hidden" name="pageURL" value="<%=Request.ServerVariables("HTTP_REFERER")%>" />
			<%=lang.getTranslated("frontend.tellafriend.label.insert_user_mail")%><br/><input type="text" name="userMail" value="" /><br/><br/>
			<%=lang.getTranslated("frontend.tellafriend.label.insert_mail_list")%><br/><input type="text" name="tellmail" value="" class="formFieldTXTLong"/><br/><br/>
			<%=lang.getTranslated("frontend.tellafriend.label.insert_mail_msg")%><br/><textarea name="tellafriendMessage" class="formFieldTXTLong"></textarea><br/><br/>
			<div align="center" style="text-align:left;">
			<%
			if(request("captcha_err") = 1) then
				response.write("<span  class=imgError>"&lang.getTranslated("frontend.tellafriend.label.wrong_captcha_code") & "</span><br/>")
			end if

			if(Application("use_recaptcha") = 0) then%>
				<br/><img id="imgCaptcha" src="<%=Application("baseroot")&"/common/include/captcha/base_captcha.asp"%>" />&nbsp;&nbsp;<input name="captchacode" type="text" id="captchacode" />
				<br/><a href="javascript:void(0)" onclick="RefreshImage('imgCaptcha')"><%'=lang.getTranslated("frontend.tellafriend.label.change_captcha_img")%></a>
			<%else%>
				<br/><%=recaptcha_challenge_writer(Application("recaptcha_pub_key"))%>
			<%end if%>
			</div>
			<input type="submit" value="submit" />
		</form>
	<%end if%>
</body>
</html>