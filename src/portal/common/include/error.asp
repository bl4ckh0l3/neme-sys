<!-- #include file="IncludeObjectList.inc" -->
<%
if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUsrLoggedEditor, objUsrLoggedEditorTmp, strRuoloEditor
	Set objUsrLoggedEditorTmp = new UserClass
	Set objUsrLoggedEditor = objUsrLoggedEditorTmp.findUserByID(Session("objCMSUtenteLogged"))
	strRuoloEditor = objUsrLoggedEditor.getRuolo()
	Set objUsrLoggedEditor = nothing
	Set objUsrLoggedEditorTmp = nothing
	
	if(strRuoloEditor = Application("admin_role") OR strRuoloEditor = Application("editor_role")) then
		response.Redirect(Application("baseroot")&"/editor/include/error.asp?error="&Server.URLEncode(request("error"))&"&id_usr="&request("id_usr"))
	end if
end if

Dim objErrorList, strErrorParam
Dim id_user
if not(request("id_usr") = "") then
	id_user = request("id_usr")
else
	id_user = -1
end if

Set errorList = Server.CreateObject("Scripting.Dictionary")
errorList.add "001", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.user_mail_already_exist")&"&nbsp;&nbsp;<a href="""& Application("baseroot") & "/area_user/manageUser.asp?id_utente="& id_user & """>"&lang.getTranslated("portal.commons.errors.label.repeat_insert")&"</a>"
errorList.add "002", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_user_found")
errorList.add "003", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.user_pwd_incorrect")
errorList.add "004", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.obj_no_found")
errorList.add "005", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_target_user")
errorList.add "006", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_target_news")
errorList.add "007", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_target_editor")
errorList.add "008", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.templates.label.page_in_progress")
errorList.add "009", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.filesystem_no_correct")
errorList.add "010", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_target_exist")		
errorList.add "011", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.target_in_use")
errorList.add "012", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.user_in_use")
errorList.add "013", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_cat_found")
errorList.add "014", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.cat_in_use")
errorList.add "015", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.lang_in_use")
errorList.add "016", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.multi_lang_in_use")
errorList.add "017", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_result_found")		
errorList.add "027", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.cat_already_exist")
errorList.add "030", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.max_content_length")
errorList.add "031", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.invalid_contenttype")
errorList.add "032", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.generic_error")
						
'****************** ERROR LIST PER ORDINI/PRODOTTI
errorList.add "018", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_product_found")
errorList.add "019", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.item_carrello_problem")
errorList.add "020", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_target_prod")
errorList.add "021", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.prod_in_use")
errorList.add "022", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.carrello_not_found")
errorList.add "023", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.user_not_enabled")
errorList.add "024", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.t_prod_or_t_lang_miss")
errorList.add "025", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.no_target_lang")
errorList.add "026", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.t_news_or_t_lang_miss")
errorList.add "028", "<span id=""error"">"&lang.getTranslated("portal.commons.errors.label.error")&":</span><br><br>"&lang.getTranslated("portal.commons.errors.label.qta_less_zero")

strErrorParam = request("error")

Response.Charset="UTF-8"
Session.CodePage  = 65001%>
<!-- #include virtual="/public/layout/include/error.asp" -->
