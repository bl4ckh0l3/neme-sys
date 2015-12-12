<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/area_user.css"%>" type="text/css">
<script language="JavaScript">
var step2ok = true;
var step3ok, step4ok = false;

function deleteUtente(){
	if(confirm("<%=lang.getTranslated("frontend.area_user.manage.label.conf_del")%>")){
		location.href = "<%=Application("baseroot") & "/area_user/deluser.asp"%>";
	}

}

function insertUser(){
	if(controllaCampiInput()){
		document.form_inserisci.submit();
	}else{
		return;
	}
}

function checkStep2(){
	step3ok, step4ok = false;
	if(document.form_inserisci.username.value == "<%=lang.getTranslated("frontend.area_user.manage.label.username")%>"){
		document.form_inserisci.username.value = "";
	}	
	if(document.form_inserisci.username.value == ""){
		alert("<%=lang.getTranslated("frontend.area_user.js.alert.insert_username")%>");
		document.form_inserisci.username.focus();
		return false;
	}

  <%if(Cint(id_utente)=-1)then%>
	/*if(document.form_inserisci.password.value == "<%=lang.getTranslated("frontend.area_user.manage.label.password")%>"){
		document.form_inserisci.password.value = "";
	}*/	
	if(document.form_inserisci.password.value == ""){
		alert("<%=lang.getTranslated("frontend.area_user.js.alert.insert_pwd")%>");
		document.form_inserisci.password.focus();
		return false;
	}	
  <%end if%>
	if(document.form_inserisci.password.value != document.form_inserisci.conferma_password.value){
		alert("<%=lang.getTranslated("frontend.area_user.js.alert.pwd_no_match")%>");
		document.form_inserisci.conferma_password.focus();
		return false;
	}

	if(document.form_inserisci.email.value == "<%=lang.getTranslated("frontend.area_user.manage.label.email")%>"){
		document.form_inserisci.email.value = "";
	}	
	var strMail = document.form_inserisci.email.value;
	if(strMail != ""){
		if (strMail.indexOf("@")<2 || strMail.indexOf(".")==-1 || strMail.indexOf(" ")!=-1 || strMail.length<6){
			alert("<%=lang.getTranslated("frontend.area_user.js.alert.wrong_mail")%>");
			document.form_inserisci.email.focus();
			return false;
		}
	}else if(strMail == ""){
		alert("<%=lang.getTranslated("frontend.area_user.js.alert.insert_mail")%>");
		document.form_inserisci.email.focus();
		return false;
	}		
	if(document.form_inserisci.email.value != document.form_inserisci.conferma_email.value){
		alert("<%=lang.getTranslated("frontend.area_user.js.alert.email_no_match")%>");
		document.form_inserisci.conferma_email.focus();
		return false;
	}

	step2ok = false;
	step3ok = true;
	return true;
}

function checkStep3(){
	step4ok = false;
	<%
	if(hasUserFields) then
	for each k in objListUserField
	  Set objField = objListUserField(k)
	  labelForm = objField.getDescription()
	  if not(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())="") then labelForm = lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())
	  response.write(objUserField.renderUserFieldJS(objField,"form_inserisci",lang,labelForm,true))
	next
	end if
	%>
	
	step2ok = false;
	step3ok = false;
	step4ok = true;
	return true;
}

function checkStep4(){
	/* DISABILITO QUESTO CONTROLLO, IL CHECKBOX DELLA NEWSLETTER NON VIENE PIU' USATO ED E' ATTIVA DI DEFAULT
	if(document.form_inserisci.ck_newsletter.checked == false){
		document.form_inserisci.newsletter.value = "false";	
	}else{
		document.form_inserisci.newsletter.value = "true";		
	}*/
	document.form_inserisci.newsletter.value = "true";

	var newsletter_values = "";
	if (document.form_inserisci.list_newsletter){
		if(document.form_inserisci.list_newsletter.length == null){
			if (document.form_inserisci.list_newsletter.checked){
				newsletter_values = newsletter_values + document.form_inserisci.list_newsletter.value + ", ";
			}
		}else{
			for (var i=0; i < document.form_inserisci.list_newsletter.length; i++){
				if (document.form_inserisci.list_newsletter[i].checked){
					newsletter_values = newsletter_values + document.form_inserisci.list_newsletter[i].value + ", ";
				}
			}
		}
		newsletter_values = newsletter_values.substring(0, newsletter_values.lastIndexOf(', '));
	}
	document.form_inserisci.list_newsletter_values.value = newsletter_values;

	step2ok = false;
	step3ok = false;
	step4ok = false;
	return true;
}

function checkStep5(){
	var strTargets = document.form_inserisci.ListTarget.value;
	if(strTargets.charAt(strTargets.length -1) == "|"){
		strTargets = strTargets.substring(0, strTargets.length -1);
	}
	document.form_inserisci.ListTarget.value = strTargets;
	
	if(document.form_inserisci.privacy.checked == false){
		alert("<%=lang.getTranslated("frontend.area_user.js.alert.privacy_confirm_needed")%>");
		document.form_inserisci.privacy.checked = true;
		return false;
	}


  <%if(Application("use_recaptcha") = 0) then%>
    // VECCHIA FUNZIONE PER CAPTCHA 	
    if(document.form_inserisci.captchacode.value == ""){
      alert("<%=lang.getTranslated("frontend.area_user.js.alert.insert_captchacode")%>");
      document.form_inserisci.captchacode.focus();
      return false;
    }
    
    // imposto campo hidden sent_captchacode 
    // perchè quello originale non viene recuperato in process
    document.form_inserisci.sent_captchacode.value = document.form_inserisci.captchacode.value;  
  <%else%>
    // FUNZIONE PER RECAPTCHA  
    if(document.form_inserisci.recaptcha_response_field.value == ""){
      alert("<%=lang.getTranslated("frontend.area_user.js.alert.insert_captchacode")%>");
      document.form_inserisci.recaptcha_response_field.focus();
      return false;
    }
      // imposto campo hidden sent_recaptcha_challenge_field e  sent_recaptcha_response_field
    // perchè quello originale non viene recuperato in process
    document.form_inserisci.sent_recaptcha_challenge_field.value = document.form_inserisci.recaptcha_challenge_field.value;
    document.form_inserisci.sent_recaptcha_response_field.value = document.form_inserisci.recaptcha_response_field.value;
  <%end if%>

	if(document.form_inserisci.del_usrimage){
		if(document.form_inserisci.del_usrimage.checked == false){
			document.form_inserisci.del_avatar.value = "false";	
		}else{
			document.form_inserisci.del_avatar.value = "true";		
		}
	}
		
	return true;
}

function controllaCampiInput(){
	<%if(Application("use_wizard_registration")="1")then%>	
	return checkStep5();
	<%else%>
	if(checkStep2() && checkStep3() && checkStep4() && checkStep5()){return true;}else{return false;}
	<%end if%>
}

function replaceChars(inString){
	var outString = inString;

	for(a = 0; a < outString.length; a++){
		if(outString.charAt(a) == '"'){
			outString=outString.substring(0,a) + "&quot;" + outString.substring(a+1, outString.length);
		}
	}
	return outString;
}

function RefreshImage(valImageId) {
	var objImage = document.images[valImageId];
	if (objImage == undefined) {
		return;
	}
	var now = new Date();
	objImage.src = objImage.src.split('?')[0] + '?x=' + now.toUTCString();
}

function checkNewsletter(formField){	
	if(document.form_inserisci.ck_newsletter.checked == false){
		formField.checked = false;
	}	
}

function uncheckNewsletter(){	
	if(document.form_inserisci.ck_newsletter.checked == false){
		if(document.form_inserisci.list_newsletter != null){
			if(document.form_inserisci.list_newsletter.length == null){
				document.form_inserisci.list_newsletter.checked = false;
			}else{
				for(i=0; i<document.form_inserisci.list_newsletter.length; i++){				
					document.form_inserisci.list_newsletter[i].checked = false;
				}
			}
		}
	}	
}

function changeTab(number){
	if(number==1)
		location.href='<%=Application("baseroot") & "/area_user/userprofile.asp"%>';
	else if(number==2)
		location.href='<%=Application("baseroot") & "/area_user/manageuser.asp"%>';
	else if(number==3)
		location.href='<%=Application("baseroot") & "/area_user/friendlist.asp"%>';
	else if(number==4)
		location.href='<%=Application("baseroot") & "/area_user/userphotos.asp"%>';

}
  
function userWizard(step){
	//var moveStep = true;

	//for(var j=2; j < step; j++){
		//eval("moveStep = moveStep && step"+j+"ok;");		 
	//}
	eval("moveStep = step"+step+"ok;");
	
	if(moveStep){
		eval("var action = checkStep"+step+"();");
		if(action==true){
			for (var i=1; i < 5; i++){
				if(i==step){
					showWizardDiv(i);
				}else{
					hideWizardDiv(i);      
				}
			}
		}
	}
}

function showWizardDiv(step){
	var element = document.getElementById('wizard'+step);
  element.style.visibility = "visible";		
  element.style.display = "block";
  
	var elementA = document.getElementById("step"+step);
  elementA.className="active";
}

function hideWizardDiv(step){
	var element = document.getElementById('wizard'+step);
  element.style.visibility = "hidden";
  element.style.display = "none";
  
  var elementA = document.getElementById("step"+step);
  elementA.className="";
}
</script>
</head>
<body>
<%if(Application("use_wizard_registration")<>"1")then%>
<script language="JavaScript">
    jQuery(document).ready(function(){
       $("#wizard2").show(); 
       $("#wizard3").show();
       $("#wizard4").show();
    });
</script>
<%end if%>
<!-- #include file="grid_top.asp" -->
		<form action="<%=Application("baseroot") & "/area_user/processUserSito.asp"%>" method="post" name="form_inserisci" enctype="multipart/form-data" accept-charset="UTF-8">		  
		<input type="hidden" value="<%=id_utente%>" name="id_utente">
		<input type="hidden" name="user_active" value="<%=bolUserActive%>">
		<input type="hidden" name="sconto" value="<%=numSconto%>">
		<input type="hidden" name="admin_comments" value="<%=strAdminComments%>">
		<input type="hidden" value="<%=dateInsertDate%>" name="insertDate">
		<input type="hidden" value="<%=dateModifyDate%>" name="modifyDate">
		<input type="hidden" name="newsletter" value="<%=bolNewsletter%>">
		<input type="hidden" name="list_newsletter_values" value="">
		<input type="hidden" name="sent_captchacode" value="">
		<input type="hidden" name="ruolo_utente" value="<%=strUsrRuolo%>">
		<input type="hidden" name="del_avatar" value="0">
		<input type="hidden" name="user_group" value="<%=numUserGroup%>">
		<input type="hidden" name="sent_recaptcha_challenge_field" value="">
		<input type="hidden" name="sent_recaptcha_response_field" value="">
		<%if not (isNull(objUsrTarget)) then
			Dim paramTarget
			paramTarget = ""
			for each y in objUsrTarget.Keys
				if(Cint(objUsrTarget(y).isAutomatic())=0) then
					paramTarget = paramTarget & y & "|"%>	
				<%end if
			next
		end if%>
		<input type="hidden" value="<%=paramTarget%>" name="ListTarget">

		
        <h1><%=lang.getTranslated("frontend.header.label.utente_modify")%>&nbsp;<em><%=strUserName%></em>
	<%if (Cint(id_utente) <> -1) then%>&nbsp;<input name="delete" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.del_user")%>" type="button" onclick="javascript:deleteUtente();"><%end if%>
	</h1>

	<!--nsys-modcommunity1-->
        <p>
	<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.profile")%>" type="button" onclick="javascript:changeTab(1);">
	<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.modify")%>" type="button" onclick="javascript:changeTab(2);">
	<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.friends")%>" type="button" onclick="javascript:changeTab(3);">
	<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.photos")%>" type="button" onclick="javascript:changeTab(4);">
	</p>
	<!---nsys-modcommunity1-->

	<p style="padding-top:10px;padding-bottom:10px;"><%=lang.getTranslated("frontend.area_user.manage.label.txt_intro_registrazione")%></p>
      
	<%if(Application("use_wizard_registration")="1")then%>
	  <div id="profilo-utente-wizard">
		<span class="active" id="step1">STEP 1</span>&nbsp;-&nbsp;<a href="javascript:userWizard(2);" id="step2">STEP 2</a>&nbsp;-&nbsp;<a href="javascript:userWizard(3);" id="step3">STEP 3</a>&nbsp;-&nbsp;<a href="javascript:userWizard(4);" id="step4">STEP 4</a>
	  </div>
	<%end if%>        
	<div id="profilo-utente">
      
        <div id="wizard1">
             <h2><%=lang.getTranslated("frontend.header.label.utente_profile_group")%></h2>
               <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.username")%> (*)</span></div>
                  <div class="vals">
          <%if (Cint(id_utente) <> -1) then%>
            <em><%=strUserName%></em><input type="hidden" name="username" id="username" value="<%=strUserName%>">				
          <%else%>
            <input type="text" name="username" id="username" value="<%=strUserName%>" onfocus="cleanInputField('username');" onBlur="restoreInputField('username','<%=lang.getTranslated("frontend.area_user.manage.label.username")%>');">
          <%end if%>
          </div>
          <div>
          <%
          If not(request.cookies(Application("srt_default_server_name"))("id_user") = "") Then%>
            <a href="<%=Application("baseroot")&"/area_user/manageuser.asp?del_autologin=1"%>"><%=lang.getTranslated("frontend.area_user.manage.label.reset_auto_login")%></a>
            <%
            if(request("del_autologin")="1") then
              response.cookies(Application("srt_default_server_name")).Expires=DateAdd("d",-1,date())
            end if
          End If
          %>
          </div>
 
	  <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.password")%> (*)</span></div>  				
            <div class="vals"><input type="password" name="password" id="password" value="" onkeypress="javascript:return notSpecialCharAndSpace(event);"/></div>
	  <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.conf_password")%> (*)</span></div> 
            <div class="vals"><input name="conferma_password" id="conferma_password" type="password" value=""  onkeypress="javascript:return notSpecialCharAndSpace(event);"/></div>	
	  
                  <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.email")%> (*)</span></div>
                  <div class="vals"><input type="text" name="email" value="<%=strEmail%>" id="email" onfocus="cleanInputField('email');" onBlur="restoreInputField('email','<%=strEmail%>');"/></div>
                  <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.confirm_email")%> (*)</span></div>
		  <div class="vals">
          <%if (Cint(id_utente) <> -1) then%>
            <input type="text" name="conferma_email" id="conferma_email" value="<%=strEmail%>" onfocus="cleanInputField('conferma_email');" onBlur="restoreInputField('conferma_email','<%=strEmail%>');"/>
          <%else%>
            <input type="text" name="conferma_email" id="conferma_email" value="<%=lang.getTranslated("frontend.area_user.manage.label.confirm_email")%>" onfocus="cleanInputField('conferma_email');" onBlur="restoreInputField('conferma_email','<%=lang.getTranslated("frontend.area_user.manage.label.confirm_email")%>');"/>
          <%end if%>
          </div>	

        <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.public_profile")%></span></div>	
        <div class="vals"><select name="public_profile" id="public_profile">
                  <OPTION VALUE="1" <%if (strComp("1", bolPublic, 1) = 0) then response.Write("selected")%>><%=lang.getTranslated("portal.commons.yes")%></OPTION>
                  <OPTION VALUE="0" <%if (strComp("0", bolPublic, 1) = 0) then response.Write("selected")%>><%=lang.getTranslated("portal.commons.no")%></OPTION>
          </select>
        </div>	

        <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.avatar")%></span></div>
        <div class="vals" id="profilo-utente-img">
        <input type="file" name="imageupload" /><br/><br/>
        <%if (usrHasImg) then%>
        <script>
          $(function() {
            $(".imgAvatarUser").aeImageResize({height: 50, width: 50});
          });
        </script>
        <img class="imgAvatarUser" align="top" src="<%=Application("baseroot") & "/common/include/userImage.asp?userID="&id_utente%>" /><input type="checkbox" align="bottom" value="false" name="del_usrimage">&nbsp;<%=lang.getTranslated("frontend.area_user.manage.label.del_avatar")%>
        <!--<script>resizeimagesByID('imgUser', 50);</script>-->
        <%end if%>

        <%
        if(request("error") = "030") then
          response.write("<span class=""imgError"">"&lang.getTranslated("portal.commons.errors.label.max_content_length")&"</span><br/>")
        end if
        if(request("error") = "031") then
          response.write("<span class=""imgError"">"&lang.getTranslated("portal.commons.errors.label.invalid_contenttype")&"</span><br/>")
        end if
        %>      
        </div>

		<%if(Application("use_wizard_registration")="1")then%>
        <div>
        <input name="wizard" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.carryon")%>" type="button" onclick="javascript:userWizard(2);">
        </div>
		<%end if%>
        </div>
              
        
        <div id="wizard2" style="display:none;">
       <!--******** GESTIONE FIELDS UTENTE PERSONALIZZATI ********-->
        <%
        '********** RECUPERO LA LISTA DI FIELD UTENTE DISPONIBILI
        Dim strPrecFieldgroup, strFieldgroup               
        strPrecFieldgroup = ""
                
        On Error Resume next
        Dim userFieldcount, fieldCssClass
        userFieldcount =1
        if(hasUserFields) then
          for each k in objListUserField
            Set objField = objListUserField(k)
            fieldCssClass=""
          
              if(CInt(objField.getTypeField())=5) then
                fieldCssClass="formFieldMultiple"
              end if
              
              labelForm = objField.getDescription()
              if not(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())="") then labelForm = lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())
          
            if(userFieldcount=1) then
              strFieldgroup = objField.getObjGroup().getDescription()
              strPrecFieldgroup = strFieldgroup%>
              <h2><%if not(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)="") then response.write(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)) else response.write(strFieldgroup) end if%></h2>
              <%if(Cint(objField.getTypeField())<>8)then%>
	      <div style="float:left;"><span><%=labelForm%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></div>
              <div class="vals"><%=objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang)%></div>
		<%else
			response.write(objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang))
		end if
	else
                strFieldgroup = objField.getObjGroup().getDescription()
                if(strFieldgroup = strPrecFieldgroup) then
			if(Cint(objField.getTypeField())<>8)then%>
				<div style="float:left;"><span><%=labelForm%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></div>
				<div class="vals"><%=objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang)%></div>
		<%	else
				response.write(objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang))
			end if
		else%>
			  <h2><%if not(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)="") then response.write(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)) else response.write(strFieldgroup) end if%></h2>
			<%if(Cint(objField.getTypeField())<>8)then%>
			  <div style="float:left;"><span><%=labelForm%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></div>
			  <div class="vals"><%=objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang)%></div>     
			<%else
				response.write(objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,lang))
			end if		  
			strPrecFieldgroup = strFieldgroup
                  end if              
              end if
                  
              'if(userFieldcount = objListUserField.Count) then response.write("</ul>") end if
                
            userFieldcount=userFieldcount+1
          next
        end if

        if(Err.number<>0) then
        'response.write(Err.description)
        end if        

      Set objListUserField = nothing
      Set objUserField = nothing
      %>
      <!--******** FINE GESTIONE FIELDS UTENTE PERSONALIZZATI ********-->

		<%if(Application("use_wizard_registration")="1")then%>            
        <div>
        <input name="wizard" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.carryon")%>" type="button" onclick="javascript:userWizard(3);">
        </div>
		<%end if%>
        </div>
        
        <div id="wizard3" style="display:none;">
        <h2><%=lang.getTranslated("frontend.header.label.iscriz_newsletter")%></h2>

        <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.iscriz_newsletter")%></span></div>
        <!--<div><input type="checkbox" value="true" name="ck_newsletter" id="ck_newsletter" <%if (bolNewsletter) then response.Write("checked")%> onclick="uncheckNewsletter();"/></div>-->	
        <div class="vals" id="profilo-utente-newsletter">
          <%
        Dim hasNewsletter, objNewsletterTmp
        hasNewsletter = false
        on error Resume Next
        
          Set objListaNewsletter = objNewsletter.getListaNewsletter(1)
          if isObject(objListaNewsletter) AND not(isNull(objListaNewsletter)) AND not (isEmpty(objListaNewsletter)) then
            if(objListaNewsletter.Count > 0) then
              hasNewsletter = true
            end if
          end if
          
        if Err.number <> 0 then
          'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
        end if	
        
        if(hasNewsletter) then
            dim chechedVal
            for each x in objListaNewsletter.Keys			
              Set objNewsletterTmp = objListaNewsletter(x)
              if not(isNull(objNewsletterUsr)) then
                chechedVal = ""
                if objNewsletterUsr.Exists(x)= true then chechedVal = "checked" end if
              end if
              %>		  
              <%=objNewsletterTmp.getDescrizione()%><input type="checkbox" value="<%=x%>" onclick="/*checkNewsletter(this);*/" align="left" name="list_newsletter" id="list_newsletter" <%=chechedVal%>><br/>
              <%Set objNewsletterTmp = nothing
            next%>
        <%end if%>
        </div>		

		<%if(Application("use_wizard_registration")="1")then%>
        <div>
        <input name="wizard" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.carryon")%>" type="button" onclick="javascript:userWizard(4);">
        </div>
		<%end if%>
        </div>
        
        <div id="wizard4" style="display:none;">
        <h2><%=lang.getTranslated("frontend.header.label.info_privacy")%></h2>

        <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.info_privacy")%></span></div>
        <div class="vals" id="profilo-utente-privacy"><textarea name="txt_privacy"><%=lang.getTranslated("frontend.area_user.manage.label.text_privacy")%></textarea></div>
	
        <div style="float:left;"><span><%=lang.getTranslated("frontend.area_user.manage.label.confirm_privacy")%> (*)</span></div>
        <div class="vals" id="profilo-utente-confirm-privacy"><input type="checkbox" value="true" name="privacy" checked/ ></div>	

        

        <div>
        <%
          if(request("captcha_err") = 1) then
            response.write("<span  class=imgError>"&lang.getTranslated("frontend.area_user.manage.label.wrong_captcha_code") & "</span><br/>")
          end if
          
          if(Application("use_recaptcha") = 0) then%>
            <br/><img id="imgCaptcha" src="<%=Application("baseroot")&"/common/include/captcha/base_captcha.asp"%>" />&nbsp;&nbsp;<input name="captchacode" type="text" id="captchacode" />
            <br/><a href="javascript:void(0)" onclick="RefreshImage('imgCaptcha')"><%=lang.getTranslated("frontend.area_user.manage.label.change_captcha_img")%></a>
          <%else%>
            <br/><%=recaptcha_challenge_writer(Application("recaptcha_pub_key"))%>
          <%end if%>
          <br/><br/>
          
          <input name="send" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.do_registration")%>" type="button" onclick="javascript:insertUser();">
	  </div>	

        </div>
       </div>
		</form>	
		   
<!-- #include file="grid_bottom.asp" -->
</body>
</html>