<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/CommentsClass.asp" -->
<%
'<!--nsys-usrins1-->
%>
<!-- #include virtual="/common/include/Objects/UserGroupClass.asp" -->
<%
'<!---nsys-usrins1-->
%>
<!-- #include virtual="/common/include/Objects/UserFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->
<!-- #include file="include/init2.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/highcharts.js"%>"></script>
<script language="JavaScript">
function insertUser(){
	if(controllaCampiInput()){
		document.form_inserisci.submit();
	}else{
		return;
	}
}

function move(fbox, tbox){
	var arrFbox = new Array();
	var arrTbox = new Array();
	var arrLookup = new Array();
	var i;
	
	for(i = 0; i < tbox.options.length; i++){
		arrLookup[tbox.options[i].text] = tbox.options[i].value;
		arrTbox[i] = tbox.options[i].text;
	}
	
	var fLength = 0;
	var tLength = arrTbox.length;
	
	for(i = 0; i < fbox.options.length; i++){
		arrLookup[fbox.options[i].text] = fbox.options[i].value;
		if(fbox.options[i].selected && fbox.options[i].value != ""){
			arrTbox[tLength] = fbox.options[i].text;
			tLength++;
		}else{
			arrFbox[fLength] = fbox.options[i].text;
			fLength++;
		}
	}
	
	arrFbox.sort();
	arrTbox.sort();
	fbox.length = 0;
	tbox.length = 0;
	var c;
	
	for(c = 0; c < arrFbox.length; c++){
		var no = new Option();
		no.value = arrLookup[arrFbox[c]];
		no.text = arrFbox[c];
		fbox[c] = no;
	}
	
	for(c = 0; c < arrTbox.length; c++){
		var no = new Option();
		no.value = arrLookup[arrTbox[c]];
		no.text = arrTbox[c];
		tbox[c] = no;
	}
}


function controllaCampiInput(){		
	//valorizzo il campo nascosto "ListTarget" con la lista dei Target della news separati da "|"
	var strTargets = "";
	strTargets+=listTargetAll
	if(strTargets.charAt(strTargets.length -1) == "|"){
		strTargets = strTargets.substring(0, strTargets.length -1);
	}
	
	document.form_inserisci.ListTarget.value = strTargets;
	//alert(document.form_inserisci.ListTarget.value);

	
	if(document.form_inserisci.username.value == ""){
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.insert_username")%>");
		document.form_inserisci.username.focus();
		return false;
	}
	
	<%if(Cint(id_utente)=-1)then%>
	if(document.form_inserisci.password.value == ""){
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.insert_pwd")%>");
		document.form_inserisci.password.focus();
		return false;
	}
	<%end if%>
	if(document.form_inserisci.password.value != document.form_inserisci.conferma_password.value){
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.pwd_no_match")%>");
		document.form_inserisci.conferma_password.focus();
		return false;
	}

	var strMail = document.form_inserisci.email.value;
	if(strMail != ""){
		if (strMail.indexOf("@")<2 || strMail.indexOf(".")==-1 || strMail.indexOf(" ")!=-1 || strMail.length<6){
			alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.wrong_mail")%>");
			document.form_inserisci.email.focus();
			return false;
		}
	}else if(strMail == ""){
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.insert_mail")%>");
		document.form_inserisci.email.focus();
		return false;
	}	
/*<!--nsys-usrins2-->*/
	if(document.form_inserisci.sconto.value != "") {
		var scontoTmp = document.form_inserisci.sconto.value;
		if(!checkDoubleFormat(scontoTmp) || scontoTmp.indexOf(".")!=-1){
			alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.isnan_value")%>");
			document.form_inserisci.sconto.value = "0";
			document.form_inserisci.sconto.focus();
			return false;
		}
	}else{
		alert("<%=langEditor.getTranslated("backend.utenti.detail.js.alert.insert_sconto")%>");
		document.form_inserisci.sconto.value = "0";
		document.form_inserisci.sconto.focus();
		return false;		
	}
/*<!---nsys-usrins2-->*/

	<%
	if(hasUserFields) then
		for each k in objListUserField
			Set objField = objListUserField(k)
			labelForm = objField.getDescription()
			if not(langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription())="") then labelForm = langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription())
			response.write(objUserField.renderUserFieldJS(objField,"form_inserisci",langEditor,labelForm,true))
		next
	end if
	%>


	if(document.form_inserisci.ck_newsletter.checked == false){
		document.form_inserisci.newsletter.value = "false";	
	}else{
		document.form_inserisci.newsletter.value = "true";		
	}
	
	return true;
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
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">
		<form action="<%=Application("baseroot") & "/editor/utenti/ProcessUtente.asp"%>" method="post" name="form_inserisci">
		  <input type="hidden" value="<%=id_utente%>" name="id_utente">
		  <input type="hidden" value="<%=dateInsertDate%>" name="insertDate">
		  <input type="hidden" value="<%=dateModifyDate%>" name="modifyDate">
<!--nsys-usrins3-->
<!---nsys-usrins3-->
			<tr>
			<td>	 		  
		    <span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.username")%></span><br>
			<%if (Cint(id_utente) <> -1) then%>
				<div><%="<b>"&strUserName &"</b>"%></div>
				<%
				' widget grafico preferenza utente
				dim percentual, total, total_comment_news, objUserPreference
				percentual = 0
				Set objUserPreference = new UserPreferenceClass
				percentual = objUserPreference.findUserPreferencePositivePercent(id_utente)
				percentual = FormatNumber(percentual, 0,-1)
				total = objUserPreference.findNumUserPreferenceTotal(id_utente, true)
				Set objComment = new CommentsClass
				total_comment_news = objComment.countCommentiByIDUtente(id_utente,1,1)
'<!--nsys-usrins4-->
                                dim total_comment_prod
				total_comment_prod = objComment.countCommentiByIDUtente(id_utente,2,1)
'<!---nsys-usrins4-->
				Set objComment = nothing
				%>
				<div style="padding-left:0px;padding-top:3px;" class="txtUserPreference">
				<%=langEditor.getTranslated("backend.utenti.detail.table.label.like")%>:&nbsp;<%=percentual%>%<br/>				
				<script type="text/javascript">
				      $(function () {
					var chart;
					$(document).ready(function() {
					  chart = new Highcharts.Chart({
					    chart: {
					      renderTo: 'usrprefchartbox',
					      type: 'bar',
					      width: 100,
					      height: 70,
					      spacingTop:-15,
					      marginLeft:-1,
					      marginRight:0
					    },
					    title: {
					      text: ''
					    },
					    xAxis: {
					      categories: [''],
					      gridLineWidth:0
					    },
					    yAxis: {
					      title: {
						text: ''
					      },
					      min: 0,
					      max:100,
					      showFirstLabel:false,
					      showLastLabel:false,
					      gridLineWidth:0
					    },
					    tooltip: {
					      enabled: false
					    },            
					    legend: {
					      enabled: false
					    },
					    series: [{
					      data: [
					      {
						color: 'blue',
						y: <%=percentual%>
					      }
					      ]
					    }]
					  });
					});
				      });
				</script>
				<div align="left" id="usrprefchartbox" style="width:100px;height:5px;border:#000000 1px solid;overflow: hidden;"></div>	
				<%=langEditor.getTranslated("backend.utenti.detail.table.label.total_vote")%>:&nbsp;<%=total%><br/>
				<%=langEditor.getTranslated("backend.utenti.detail.table.label.total_commenti_news")%>:&nbsp;<%=total_comment_news%>
<!--nsys-usrins5-->
                                <br/><%=langEditor.getTranslated("backend.utenti.detail.table.label.total_commenti_prod")%>:&nbsp;<%=total_comment_prod%>
<!---nsys-usrins5-->
				<div><%=langEditor.getTranslated("backend.utenti.detail.table.label.last_modify")&": "&FormatDateTime(dateModifyDate,2)%></div>
				</div>
				<%
				Set objUserPreference = nothing
				%>
				<input type="hidden" name="username" value="<%=strUserName%>">
			<%else%>
				<input type="text" name="username" value="<%=strUserName%>" class="formFieldTXT">
			<%end if%>
			</td>
			<td>								
			<span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.password")%></span><br>
			<input type="password" name="password" value="<%'=strPwd%>" class="formFieldTXT" onkeypress="javascript:return notSpecialCharAndSpace(event);">
			<br>
			<span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.conf_password")%></span><br>
			<input type="password" name="conferma_password" value="<%'=strPwd%>" class="formFieldTXT" onkeypress="javascript:return notSpecialCharAndSpace(event);">
<!--nsys-usrins6-->
                      <br/><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.user_group")%></span><br>
		      <select name="user_group" class="formFieldTXT">
			<option value=""></option>
			<%
			if (Instr(1, typename(objDispGroup), "dictionary", 1) > 0) then
			for each x in objDispGroup%>
			<option value="<%=x%>" <%if (numUserGroup = x) then response.Write("selected")%>><%=objDispGroup(x).getShortDesc()%></option>
			<%next
			end if%>
		      </select>
<!--nsys-usrins6-->
			</td>
			<td><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.user_role")%></span><br>
              <select name="ruolo_utente" class="formFieldTXT">
                <%for each x in objListaRuoli.Keys%>
                <option value="<%=x%>" <%if (strUsrRuolo = x) then response.Write("selected")%>><%=objListaRuoli(x)%></option>
                <%next%>
              </select>
	      <br><br>
              <span class="labelForm">
              <%=langEditor.getTranslated("backend.utenti.detail.table.label.user_active")%>&nbsp;
              <select name="user_active" class="formFieldSelectSimple">
                <OPTION VALUE="0" <%if (strComp("0", bolUserActive, 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
                <OPTION VALUE="1" <%if (strComp("1", bolUserActive, 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
              </select>
              </span>
	      <br><br>	      
              <span class="labelForm">
              <%=langEditor.getTranslated("backend.utenti.detail.table.label.public_profile")%>&nbsp;
              <select name="public_profile" class="formFieldSelectSimple">
                <OPTION VALUE="1" <%if (strComp("1", bolPublic, 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.yes")%></OPTION>
                <OPTION VALUE="0" <%if (strComp("0", bolPublic, 1) = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.commons.no")%></OPTION>
              </select>
              </span>		
		</td>
		</tr>

		<tr>
		<td>			
		<span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.email")%></span><br>
		<input type="text" name="email" value="<%=strEmail%>" class="formFieldTXT">
		</td>
		<td>
<!--nsys-usrins7-->
		<br><span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.sconto")%></span>&nbsp;
		<input type="text" name="sconto" value="<%=numSconto%>" class="formFieldTXTShort" maxlength="5" onkeypress="javascript:return isDouble(event);">
		  % 
<!---nsys-usrins7-->
		</td>
		<td>&nbsp;</td>
		</tr>

		<tr>
		<td colspan="3">&nbsp;</td>
		</tr>



		<tr>
		<td colspan="3">
		<%
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
				if not(langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription())="") then labelForm = langEditor.getTranslated("backend.utenti.detail.table.label."&objField.getDescription())
				%>
				<div align="left" style="display:inline-block;text-align:left;vertical-align:top;padding-right:10px;min-width:250px;height:60px; <%if not((userFieldcount Mod 3) = 0) then response.write("float:left;")%> <%if(Cint(objField.getTypeField())=5)then response.write("padding-bottom:20px;")%>">
				<span class="labelForm"><%=labelForm%> <%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span><br>
				<%if(Cint(objField.getTypeField())<>8)then%>
					<%=objUserField.renderUserFieldHTML(objField,fieldCssClass, id_utente, labelForm,langEditor)%>
				<%else
					fieldMatchValue=""
					if not(id_utente="") then
						on error resume next
						Set fieldMatchValue = objUserField.findFieldMatch(objField.getID(),id_utente)
						if (Instr(1, typename(fieldMatchValue), "dictionary", 1) > 0) then
							fieldMatchValue = fieldMatchValue.Item("value")
						end if
						if Err.number <> 0 then
							'response.write(Err.description)
							fieldMatchValue=""
						end if			
					end if%>
					<input type="text" value="<%=fieldMatchValue%>" name="userfield<%=objField.getID()%>" id="userfield<%=objField.getID()%>"/>
				<%end if%>
				</div>
			<%userFieldcount=userFieldcount+1
			next
		end if

		Set objListUserField = nothing
		Set objUserField = nothing

		if(Err.number<>0) then
		'response.write(Err.description)
		end if
		%>
		</td>
		</tr>

			
			<tr>		
			<td colspan="3" align="left" nowrap>	
			  <br>
			  <br>
		<input type="hidden" value="" name="ListTarget">
		<%
		Set objT = New TargetClass
		response.write(objT.renderTargetBox("listTargetAll", "targetcatbox_sx","targetcatbox_dx",langEditor.getTranslated("backend.utenti.detail.table.label.target_x_user"), langEditor.getTranslated("backend.utenti.detail.table.label.target_disp"), "1,2,3", objUsrTarget, objListaTarget, true, true, langEditor))
		Set objT = Nothing
		%>		
			  
			  <br>
			</td>
			</tr>
			<tr>		
			<td valign="top">
			  <span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.admin_comments")%></span><br>
              <textarea name="admin_comments" class="formFieldAdminComments"><%=strAdminComments%></textarea>
			</td>
			<td colspan="2">	
			  <input type="checkbox" value="true" name="ck_newsletter" <%if (bolNewsletter) then response.Write("checked")%> onclick="uncheckNewsletter();">&nbsp;<span class="labelForm"><%=langEditor.getTranslated("backend.utenti.detail.table.label.subscribe_newsletter")%></span>
              <br>
              <input type="hidden" name="newsletter" value="">
			<input type="hidden" name="privacy" value="true">
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
							<input type="checkbox" value="<%=x%>" onclick="checkNewsletter(this);" name="list_newsletter" <%=chechedVal%>>&nbsp;<%=objNewsletterTmp.getDescrizione()%><br/>				  
							<%Set objNewsletterTmp = nothing
						next%>
				<%end if
				
				Set objListaTarget = nothing
				Set objListaRuoli = nothing
				Set objSelUtente = nothing
				Set objNewsletter = nothing				
				%>
			</td>
		  </tr>
		</form>				
		</table><br/>			    
		  <input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.utenti.detail.button.inserisci.label")%>" onclick="javascript:insertUser();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/utenti/ListaUtenti.asp?cssClass=LU"%>';" />
		<br/><br/>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>