<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
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

if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objUserLoggedTmp = nothing
	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("guest_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if

	Dim field_name, field_val, objtype, id_objref
	field_name = request("field_name")
	field_val = request("field_val")
	objtype = request("objtype")
	id_objref = request("id_objref")

	Dim objRef, objTmp, objDict
	Select Case objtype
		Case "content"
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Set objRef = New NewsClass
			Set objTmp = objRef.findNewsByID(id_objref)
			objDict.add "titolo",  objTmp.getTitolo()
			objDict.add "abstract1", objTmp.getAbstract1()
			objDict.add "abstract2", objTmp.getAbstract2()
			objDict.add "abstract3", objTmp.getAbstract3()
			objDict.add "testo", objTmp.getTesto()
			objDict.add "keyword", objTmp.getKeyword()
			objDict.add "news_data", objTmp.getDataInsNews()
			objDict.add "news_data_pub", objTmp.getDataPubNews()
			objDict.add "news_data_del", objTmp.getDataDelNews()
			objDict.add "stato_news", objTmp.getStato()
			objDict.add "meta_description", objTmp.getMetaDescription()
			objDict.add "meta_keyword", objTmp.getMetaKeyword()
			objDict.add "page_title", objTmp.getPageTitle()			
			Set objTmp = nothing

			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"		
			'objDict.remove(field_name)
			'objDict.add field_name, field_val
			objDict.Item(field_name) = field_val

			'patch per formato date
			news_data_pub = objDict.item("news_data_pub")
			news_data_pub = convertDate(news_data_pub)
			' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
			'objDict.remove("news_data_pub")
			'objDict.add "news_data_pub", news_data_pub
			objDict.Item("news_data_pub") = news_data_pub

			call objRef.modifyNewsNoTransaction(id_objref, objDict.item("titolo"), objDict.item("abstract1"), objDict.item("abstract2"), objDict.item("abstract3"), objDict.item("testo"), objDict.item("keyword"), objDict.item("news_data"), objDict.item("news_data_pub"), objDict.item("news_data_del"), objDict.item("stato_news"), objDict.item("meta_description"), objDict.item("meta_keyword"), objDict.item("page_title"))
			Set objRef = nothing
			Set objDict = nothing
		Case Else			
	End Select
else
	response.Redirect(Application("baseroot")&"/login.asp")
end if
%>
