<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=langEditor.getTranslated("frontend.error_page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="<%=Application("baseroot") & "/editor/css/stile.css"%>" type="text/css">
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->		
		<div id="backend-content">
		<p class="error-text">
		<%
		Dim objErrorList, strErrorParam
		Dim id_user
		if not(request("id_usr") = "") then
			id_user = request("id_usr")
		else
			id_user = -1
		end if
		
		Set errorList = Server.CreateObject("Scripting.Dictionary")
		errorList.add "001", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.user_mail_already_exist")&"&nbsp;&nbsp;<a href="""& Application("baseroot") & "/editor/utenti/InserisciUtente.asp?id_utente="& id_user & """>"&langEditor.getTranslated("portal.commons.errors.label.repeat_insert")&"</a>"
		errorList.add "002", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_user_found")
		errorList.add "003", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.user_pwd_incorrect")
		errorList.add "004", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.obj_no_found")
		errorList.add "005", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_target_user")
		errorList.add "006", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_target_news")
		errorList.add "007", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_target_editor")
		errorList.add "008", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.templates.label.page_in_progress")
		errorList.add "009", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.filesystem_no_correct")
		errorList.add "010", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_target_exist")		
		errorList.add "011", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.target_in_use")
		errorList.add "012", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.user_in_use")
		errorList.add "013", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_cat_found")
		errorList.add "014", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.cat_in_use")
		errorList.add "015", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.lang_in_use")
		errorList.add "016", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.multi_lang_in_use")
		errorList.add "017", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_result_found")		
		errorList.add "027", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.cat_already_exist")
		errorList.add "028", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.newsletter_in_use")
		errorList.add "029", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.user_already_in_use")
		errorList.add "030", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.max_content_length")
		errorList.add "031", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.invalid_contenttype")
		errorList.add "032", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.generic_error")
									
		'****************** ERROR LIST PER ORDINI/PRODOTTI
		errorList.add "018", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_product_found")
		errorList.add "019", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.item_carrello_problem")
		errorList.add "020", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_target_prod")
		errorList.add "021", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.prod_in_use")
		errorList.add "022", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.carrello_not_found")
		errorList.add "023", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.user_not_enabled")
		errorList.add "024", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.t_prod_or_t_lang_miss")
		errorList.add "025", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.no_target_lang")
		errorList.add "026", "<span class=""labelForm"">"&langEditor.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&langEditor.getTranslated("portal.commons.errors.label.t_news_or_t_lang_miss")
	
		
		strErrorParam = request("error")
		if(errorList.Exists(Trim(LCase(strErrorParam)))) then
			response.write(errorList.Item(strErrorParam))
		else
			response.Write(strErrorParam)
		end if
		Set errorList = Nothing
		%>
		</p>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>
