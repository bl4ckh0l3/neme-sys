<%

	Public Function getAmount(amount,margin,discount,applyProdDisc,applyUserDisc,prodDiscount,userDiscount)
		'*************** gestione della logica di calcolo dei margini/sconti
		'*************** in associazione con lo sconto per il cliente e lo sconto per il prodotto

		'*************** verifico se devono essere applicati gli sconti prodotto e cliente, e li aggiungo allo sconto esistente
		if(applyProdDisc=1)then
			discount = discount+prodDiscount
		end if
		if(applyUserDisc=1)then
			discount = discount+userDiscount
		end if

		margin = margin-discount

		getAmount = amount + (amount / 100 * CDbl(margin))	
		


		'	dim objTassa, importo, iValore

		'	Set objTassa = new TaxsClass
			
		'	iValore = objTassa.findTassaByID(idTassaApplicata).getValore()
		'	iValore = CDbl(iValore)
		'	if(objTassa.findTassaByID(idTassaApplicata).getTipoValore() = 2) then
		'		importo = dblPrezzo * (iValore / 100)
		'	else
		'		importo = iValore
		'	end if
			
		'	getImportoTassa = importo
		'	Set objTassa = nothing
	End Function
	
	
	'amount = 10
	'margin = 3.5
	'discount = 1.5
	'bolProd = 0
	'bolUser=0
	'prodD= 5
	'userD=3
	
	'response.write("getAmount: "&getAmount(amount,margin,discount,bolProd,bolUser,prodD,userD))
	
	
	
	'dim strProva
	'strProva = null
	
	'strProva = Split(Trim(strProva), ", ", -1, 1)
	'strProva = join(strProva)	
	
	'response.write("<br/><br/>strProva: "&strProva)



	'Dim url, objHttp

	'On Error Resume Next
	'url = "http://"&Application("srt_default_server_name")&Application("baseroot")&"/editor/currency/currencyPoller.asp"			
	'set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
	'objHttp.open "POST", url, false
	'objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
	'objHttp.Send("test=1&prova=2")
	'Set objXML = objHTTP.ResponseXML
	'set items = objXML.getElementsByTagName("currency_confirmed")
	'val = items(0).childNodes(0).nodeValue
	'response.write(objHTTP.responseText)
	'if(Err.number <> 0) then
	'	response.write("Error: "&Err.number & " - " & Err.description)
	'else
	'	response.write("typename(items): "&typename(items))
	'	response.write(" val: "&val)
	'end if
	'set objHttp = nothing


	Public Sub checkModuleTag
		
	End Sub

	Public Sub ricorsiveFolder(objF)
		Response.write("<b>Curr fold Name:</b> "&objF.Name & "<b> - Path:</b> "&objF.Path&"<br />")
		for each y in objF.files
			'Response.write("<b>Filter Name:</b> "&Right(y.Name,(Len(y.Name)-InStrRev(y.Name,".",-1,1)))&"<br />")
			if("asp"=Right(y.Name,(Len(y.Name)-InStrRev(y.Name,".",-1,1))) OR "inc"=Right(y.Name,(Len(y.Name)-InStrRev(y.Name,".",-1,1))) OR "css"=Right(y.Name,(Len(y.Name)-InStrRev(y.Name,".",-1,1))) OR "js"=Right(y.Name,(Len(y.Name)-InStrRev(y.Name,".",-1,1))))then
				Response.write(y.Name & "<br />")
				
			end if
		next	
		if(objF.SubFolders.count>0)then
			for each x in objF.SubFolders
				ricorsiveFolder(x)
			next
		end if
	End Sub

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	response.write("start: "&now()&"<br><br>")
	set objFO=objFSO.GetFolder(Server.MapPath(Application("baseroot")&"/area_user"))
	ricorsiveFolder(objFO)
	set objFO=objFSO.GetFolder(Server.MapPath(Application("baseroot")&"/common"))
	ricorsiveFolder(objFO)
	set objFO=objFSO.GetFolder(Server.MapPath(Application("baseroot")&"/editor"))
	ricorsiveFolder(objFO)
	set objFO=objFSO.GetFolder(Server.MapPath(Application("baseroot")&"/public/layout"))
	ricorsiveFolder(objFO)
	set objFO=nothing
	response.write("end: "&now()&"<br><br>")	
	Set objFSO = nothing
%>