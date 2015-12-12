<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!--nsys-nwsletins1-->
<!-- #include virtual="/common/include/Objects/VoucherClass.asp" -->
<!---nsys-nwsletins1-->
<!-- #include file="include/init2.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
function insertNewsletter(){
	
	if(document.form_inserisci.descrizione.value == ""){
		alert("<%=langEditor.getTranslated("backend.newsletters.detail.js.alert.insert_newsletter_value")%>");
		document.form_inserisci.reset();
		document.form_inserisci.descrizione.focus();
		return;
	}
		
	document.form_inserisci.submit()
}
</script>
</head>
<body onLoad="javascript:document.form_inserisci.descrizione.focus();">
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<form action="<%=Application("baseroot") & "/editor/newsletter/ProcessNewsletter.asp"%>" method="post" name="form_inserisci">
		<input type="hidden" value="<%=id_newsletter%>" name="id_newsletter">
<!--nsys-nwsletins2-->
<!---nsys-nwsletins2-->
		<table border="0" cellspacing="0" cellpadding="0" class="principal">
			<tr>
			<td>
			<div style="float:left;padding-right:20px;">
			<span class="labelForm"><%=langEditor.getTranslated("backend.newsletters.detail.table.label.descrizione")%></span><br>
			  <input type="text" name="descrizione" value="<%=strDescrizione%>" class="formFieldTXT">
			</div>
			<div style="display:block;">		
			  <span class="labelForm"><%=langEditor.getTranslated("backend.newsletter.detail.table.header.newsletter_stato")%></span><br>
				<select name="stato" class="formFieldTXT">
				<option value="0"<%if (0=Cint(iStato)) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.newsletter.lista.table.label.inactive")%></option>	
				<option value="1"<%if (1=Cint(iStato)) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.newsletter.lista.table.label.active")%></option>	
				</SELECT>		  
			</div>
			<div style="display:block; text-align:left;margin-top:20px;margin-bottom:20px;"> 
			  <span class="labelForm"><%=langEditor.getTranslated("backend.newsletter.detail.table.header.newsletter_template")%></span><br>
				<select name="template" class="formFieldTXT">		  
				<%For y=LBound(listTemplate) to UBound(listTemplate)%>					
				<option value="<%=response.Write(listTemplate(y))%>"<%if (listTemplate(y)=strTemplate) then response.Write(" selected")%>><%=response.Write(listTemplate(y))%></option>	
				<%Next%>
				</SELECT>		  
			 </div>
<!--nsys-nwsletins3-->			 
			 <%
			Set objVoucherClass =  new VoucherClass				
			On Error Resume Next
			hasVoucherCampaign = false
			Set objListVoucherCampaign = objVoucherClass.getCampaignList(4, 1)
			if(objListVoucherCampaign.count>0)then
				hasVoucherCampaign = true
			end if
			if(Err.number <> 0)then
				hasVoucherCampaign = false
			end if
			Set objVoucherClass = nothing			 
			 %>
			<div style="display:block; text-align:left;margin-top:20px;margin-bottom:20px;"> 
			  <span class="labelForm"><%=langEditor.getTranslated("backend.newsletter.detail.table.header.voucher_campaign")%></span><br>
				<select name="voucher" class="formFieldTXT">		  
				  <option value=""></option>
				  <%
				  if(hasVoucherCampaign)then
					for each g in objListVoucherCampaign%>
					<option value="<%=g%>" <%if(g=id_voucher_campaign)then response.write(" selected") end if%>><%=objListVoucherCampaign(g).getLabel()%></option>
					<%next
				  end if
				  %>
				</SELECT>	  
			 </div>
<!---nsys-nwsletins3-->
			</td>
			</tr>
		</table>		
			<br/>				
		  <input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.newsletter.detail.button.inserisci.label")%>" onclick="javascript:insertNewsletter();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/newsletter/ListaNewsletter.asp?cssClass=LNL"%>';" />
		</form>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>