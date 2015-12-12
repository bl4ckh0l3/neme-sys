<%
'<!--nsys-editinc5-->
%>
<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ProductFieldClass.asp" -->
<!-- #include virtual="/common/include/Objects/CardClass.asp" -->
<!-- #include virtual="/common/include/Objects/NewsClass.asp" -->
<!-- #include virtual="/common/include/Objects/NewsletterClass.asp" -->
<%
'<!---nsys-editinc5-->
%>
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/common/include/Objects/UserFieldClass.asp" -->
<%
Public Function convertDate(dateToConvert)
	Dim DD, MM, YY, HH, MIN, SS
	
	convertDate = null
	
	DD = DatePart("d", dateToConvert)
	MM = DatePart("m", dateToConvert)
	YY = DatePart("yyyy", dateToConvert)
	HH = DatePart("h", dateToConvert)
	MIN = DatePart("n", dateToConvert)
	SS = DatePart("s", dateToConvert)
	
	convertDate = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS		
End Function

if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then

		Dim objLogger
		Set objLogger = New LogClass

		Dim objtype, id_objref
		objtype = request("objtype")
		id_objref = request("id_objref")

		On Error Resume Next
		Dim objRef, objTmp, objDict
		Select Case objtype
			Case "newsletter"
				Set objRef = New NewsletterClass
				if(objRef.findNewsletterAssociations(id_objref)) then	
					response.write("err:028")					
				else
					call objRef.deleteNewsletter(id_objref)
					call objLogger.write("newsletter deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				end if
				Set objRef = nothing
			Case "content"
				Set objRef = New NewsClass
				call objRef.deleteNews(id_objref)				
				call objLogger.write("content deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing

				'rimuovo l'oggetto dalla cache
				Set objCacheClass = new CacheClass
				call objCacheClass.remove("content-"&id_objref)
				call objCacheClass.remove("listcf-"&id_objref)
				call objCacheClass.removeByPrefix("findc", id_objref)
				Set objCacheClass = nothing
			Case "content_field"
				Set objRef = New ContentFieldClass
				call objRef.deleteContentField(id_objref)
				call objLogger.write("content field deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
	
				'rimuovo l'oggetto dalla cache
				Set objCacheClass = new CacheClass
				call objCacheClass.removeByPrefix("listcf-", null)
				Set objCacheClass = nothing
			Case "user"
				Set objRef = New UserClass

				Set objRef = nothing
			Case "user_field"
				Set objRef = New UserFieldClass

				Set objRef = nothing
			Case "target"
				Set objRef = New TargetClass
				if(objRef.findTargetAssociations(id_objref)) then
					response.write("err:011")		
				else
					call objRef.deleteTarget(id_objref)
					call objLogger.write("target deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")	
				end if
				Set objRef = nothing
			Case "category"
				Set objRef = New CategoryClass
				if(objRef.findCategoriaAssociations(id_objref)) then
					response.write("err:014")		
				else
					call objRef.deleteCategoria(id_objref)
					call objLogger.write("category deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")	
				end if				
				Set objRef = nothing
			Case "template"
				Set objRef = New TemplateClass
				Set objTmpTempl = objRef.findTemplateByID(id_objref)
				templateDirVar = objTmpTempl.getDirTemplate()
				templateDirVar = Application("baseroot") & Application("dir_upload_templ")& templateDirVar
				templateDirVar = Server.MapPath(templateDirVar)
				
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				if objFSO.FolderExists(templateDirVar) then
					call objFSO.DeleteFolder(templateDirVar, true)	
				end if
				Set objFSO = nothing
				call objRef.deleteTemplate(id_objref)
				call objLogger.write("template deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objTmpTempl = nothing
				Set objRef = nothing
			Case "country"
				Set objRef = New CountryClass
				call objRef.deleteCountry(id_objref)
				call objLogger.write("country deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
'<!--nsys-editinc6-->
			Case "payment"
				Set objRef = New PaymentClass
				call objRef.deletePayment(id_objref)
				call objLogger.write("payment deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "currency"
				Set objRef = New CurrencyClass
				call objRef.deleteCurrency(id_objref)
				call objLogger.write("currency deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "tax"
				Set objRef = New TaxsClass

				Set objRef = nothing
			Case "taxs_group"		
				Set objRef = New TaxsGroupClass		

				Set objRef = nothing
			Case "taxs_group_value"		
				Set objRef = New TaxsGroupClass		

				Set objRef = nothing
			Case "bill"
				Set objRef = New BillsClass
				call objRef.deleteSpesa(id_objref)
				call objLogger.write("bill deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "margin"
				Set objRef = New MarginDiscountClass
				call objRef.deleteMarginDiscount(id_objref)
				call objLogger.write("margin deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "margin_group"
				Set objRef = New UserGroupClass
				call objRef.deleteUserGroup(id_objref)
				call objLogger.write("margin_group deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "business_rule"
				Set objRef = New BusinessRulesClass
				call objRef.deleteRule(id_objref)
				call objLogger.write("business_rule deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "product"
				Set objRef = New ProductsClass
				Set objField = New ProductFieldClass
		
				Set objDB = New DBManagerClass
				Set objConn = objDB.openConnection()
				objConn.BeginTrans
				
				call objRef.deleteProdotto(id_objref,objConn)
				call objField.deleteFieldMatchByProd(id_objref, objConn)
				
				if objConn.Errors.Count = 0 then
					objConn.CommitTrans
				else
					objConn.RollBackTrans
					response.write("err:"&Err.description)
				end if				
				Set objDB = nothing
		
				call objLogger.write("product deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")	
				Set objField = nothing		
				Set objRef = nothing

				'rimuovo l'oggetto dalla cache
				Set objCacheClass = new CacheClass
				Set objBase64 = new Base64Class
				objCacheClass.remove("product-"&objBase64.Base64Encode(id_objref))
				call objCacheClass.removeByPrefix("findp", id_objref)
				Set objBase64 = nothing
				Set objCacheClass = nothing
			Case "product_field"
				Set objRef = New ProductFieldClass
				call objRef.deleteProductField(id_objref)
				call objLogger.write("product field deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "shopping_card"
				Set objRef = New CardClass	
				call objRef.deleteCarrello(id_objref)
				call objLogger.write("shopping card deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "voucher"
				Set objRef = New VoucherClass	
				call objRef.deleteCampaign(id_objref)
				call objLogger.write("voucher campaign deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
			Case "voucher_code"
				Set objRef = New VoucherClass	
				call objRef.deleteVoucherCodeNoTransaction(id_objref)
				call objLogger.write("voucher code deleted --> id: "&id_objref, objUserLogged.getUserName(), "info")
				Set objRef = nothing
'<!---nsys-editinc6-->
			Case Else			
		End Select
		
		if(Err.number<>0) then
			response.write("err:"&err.description)
		end if
		
		Set objLogger = nothing
	end if
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>
