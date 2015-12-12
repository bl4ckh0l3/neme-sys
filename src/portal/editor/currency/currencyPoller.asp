<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<%
Dim url, objXMLString, t_node, m_node, m_nodelist ,currencyAttribute, rateAttribute, dta_ins, time_ins, dta_refer

url = "http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"
		
Set objXML = Server.CreateObject("Msxml2.DOMDocument.3.0")
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
objHttp.open "POST", url, false
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send()

'*********************************** DALLA STRINGA XML ESTRAGGO SOLO I NODI CON I CAMBI	
objXMLString = objHTTP.ResponseXML.xml
objXMLString = Mid(objXMLString,InStr(1,objXMLString,"<Cube>",1),Len(objXMLString))
objXMLString = Mid(objXMLString,1,InStrRev(objXMLString,"</Cube>",-1,1)+7)

objXML.loadXML(objXMLString)

'Get the time of the xml file
Set t_node = objXML.SelectSingleNode("/Cube/Cube")

dta_refer = t_node.getAttribute("time")		
DD = DatePart("d", dta_refer)
MM = DatePart("m", dta_refer)
YY = DatePart("yyyy", dta_refer)
dta_refer = YY&"-"&MM&"-"&DD

dta_ins = Now()
DD = DatePart("d", dta_ins)
MM = DatePart("m", dta_ins)
YY = DatePart("yyyy", dta_ins)
HH = DatePart("h", dta_ins)
MIN = DatePart("n", dta_ins)
SS = DatePart("s", dta_ins)	
dta_ins = YY&"-"&MM&"-"&DD&" "&HH&":"&MIN&":"&SS

'response.write("dta_ins: "&dta_ins&"<br/>")

'Get the list of name nodes 
Set m_nodelist = objXML.SelectNodes("/Cube/Cube/Cube")


Dim objCurrencyClass, objCurrency
Set objCurrencyClass = new CurrencyClass

'Loop through the nodes
For Each m_node In m_nodelist
	
	'Get the CURRENCY Attribute Value
	currencyAttribute = m_node.getAttribute("currency")	
	'Get the RATE Attribute Value
	rateAttribute = m_node.getAttribute("rate")
	
	' faccio le verifiche se la valuta corrente esiste già sul DB, in tal caso vado in update
	On Error Resume Next
	Set objCurrency = objCurrencyClass.findCurrencyByCurrency(currencyAttribute)
	
	'response.write("objCurrency: "&typename(objCurrency)&"<br/>")
	
	if(strComp(typename(objCurrency), "CurrencyClass", 1) = 0) then
		call objCurrencyClass.modifyCurrency(objCurrency.getID(), currencyAttribute, rateAttribute, dta_refer, dta_ins, objCurrency.getActive(), objCurrency.getDefault())
	else	
		call objCurrencyClass.insertCurrency(currencyAttribute, rateAttribute, dta_refer, dta_ins, 0, 0)
	end if
	
	Set objCurrency = nothing
	
	if(Err.number <> 0) then
		'response.write(Err.description)
		response.write("<currency_confirmed>"&Err.description&"</currency_confirmed>")
	end if
	
	'response.write("currencyAttribute: "&currencyAttribute&" --> ")
	'response.write("rateAttribute: "&rateAttribute&"<br/>")
Next

response.write("<currency_confirmed>valori currency modificati</currency_confirmed>")

Set objCurrencyClass = nothing
Set objXML = nothing
set objHttp = nothing	

%>
