<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<%if not(isNull(strCSS)) ANd not(strCSS = "") then%>
<link rel="stylesheet" href="<%=Application("baseroot") & strCSS%>" type="text/css">
<%end if%>
<script type="text/JavaScript" src="<%=Application("baseroot") & "/common/js/jsDraw2D_Uncompressed.js"%>"></script> 
<script type="text/javascript" language="JavaScript">
<!--

function sendMail(){
	if(controllaCampiMail()){
		return true;
		//document.form_send_mail.submit();
	}else{
		return false;
	}
}

function controllaCampiMail(){	
	
	if(document.form_send_mail.nome.value == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_nome")%>");
		document.form_send_mail.nome.focus();
		return false;
	}	
	
	if(document.form_send_mail.cognome.value == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_cognome")%>");
		document.form_send_mail.cognome.focus();
		return false;
	}
	
	var strMail = document.form_send_mail.email.value;
	if(strMail != ""){
		if (strMail.indexOf("@")<2 || strMail.indexOf(".")==-1 || strMail.indexOf(" ")!=-1 || strMail.length<6){
			alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.alert.wrong_mail")%>");
			document.form_send_mail.email.focus();
			return false;
		}
	}else if(strMail == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_mail")%>");
		document.form_send_mail.email.focus();
		return false;
	}	
	
	if(document.form_send_mail.telefono.value == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_telefono")%>");
		document.form_send_mail.telefono.focus();
		return false;
	}		
	
	if(document.form_send_mail.nazione.value == ""){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_country")%>");
		document.form_send_mail.nazione.focus();
		return false;
	}		
	
	if(!document.form_send_mail.acceptPrivacy.checked){
		alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.confirm_privacy")%>");
		return false;
	}	

  <%if(Application("use_recaptcha") = 0) then%>
    if(document.form_send_mail.captchacode.value == ""){
      alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_captchacode")%>");
      document.form_send_mail.captchacode.focus();
      return false;
    }
    
    // imposto campo hidden sent_captchacode 
    // perchè quello originale non viene recuperato in process
    document.form_send_mail.sent_captchacode.value = document.form_send_mail.captchacode.value;
  <%else%>
    // FUNZIONE PER RECAPTCHA  
    if(document.form_send_mail.recaptcha_response_field.value == ""){
      alert("<%=lang.getTranslated("frontend.template_contatti.js.alert.insert_captchacode")%>");
      document.form_send_mail.recaptcha_response_field.focus();
      return false;
    }
      // imposto campo hidden sent_recaptcha_challenge_field e  sent_recaptcha_response_field
    // perchè quello originale non viene recuperato in process
    document.form_send_mail.sent_recaptcha_challenge_field.value = document.form_send_mail.recaptcha_challenge_field.value;
    document.form_send_mail.sent_recaptcha_response_field.value = document.form_send_mail.recaptcha_response_field.value;
  <%end if%>
	
	return true;
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