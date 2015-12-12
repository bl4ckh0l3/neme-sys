<%
'******************************* CREO UNA INNER CLASS PER GENERARE UNA COLLECTION DA UTILIZZARE NELL'ELABORAZIONE DELLA LISTA SMS
Class SmsVO
	Public cellnum
	Public message
	Public mitt

	Private Sub Class_Initialize()				
		cellnum = ""			
		message = ""			
		mitt = ""
	End Sub

	Private Sub Class_Terminate()
	End Sub
	
	Public Function toString
		'convertire il toString nel formato corretto da inviare al gestore sms
		toString = "cellnum: "&cellnum&" - message: "&message&" - mitt: "&mitt
	End Function
End Class


Dim fs,fo,objXML,secKeyConst,objHttp

secKeyConst="fsret54667hsfgqKDYT4356FSDGFDG" 'DA SPOSTARE COME VARIABILE DI TIPO APPLICATION NEL GLOBAL.ASA

Set objDictSms = Server.CreateObject("Scripting.Dictionary")
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set fo=fs.GetFolder(Server.MapPath("/public/xml4sms"))
Set objXML = Server.CreateObject("Msxml2.DOMDocument.3.0")
objXML.async = False
objXML.setProperty "SelectionLanguage", "XPath"

'verifico l'ora corrente e se è maggiore delle 22 e minore delle 8 del mattino non faccio nulla
currTime = Now()
HH = DatePart("h", currTime)
'Response.write(HH & "<br />")

if(HH >8 AND HH <22) then
	for each x in fo.files
		'Print the name of all files in the test folder
		'Response.write(x.Name & "<br />")
		'Response.write(x.Path & "<br />")
		objXML.load(x.Path)

		If (objXML.parseError.errorCode = 0) Then
			Set t_node = objXML.selectSingleNode("//ROOT/SECUREKEY")
			securekey = t_node.text
			'response.write("securekey: "&securekey&"<br><br>")
			Set t_node = nothing

			if(secKeyConst=securekey) then
				Set m_nodelist = objXML.SelectNodes("//ROOT/SMS")	
				For Each m_node In m_nodelist
					If m_node.hasChildNodes Then
						Set objSms = new SmsVO
						For Each c_node In m_node.childNodes
							if(c_node.NodeName="NUMEROCELL") then
								objSms.cellnum = c_node.text
								'response.write("cellnum: "&cellnum&"<br>")
							elseif(c_node.NodeName="MESSAGGIO") then
								objSms.message = c_node.text
								'response.write("message: "&message&"<br>")
							elseif(c_node.NodeName="MITTENTE") then
								objSms.mitt = c_node.text
								'response.write("mitt: "&mitt&"<br>") 
							end if				
						next
						
						objDictSms.add objSms,""
						Set objSms = nothing
					End If
				Next
				Set m_nodelist = nothing			
			end if
		End If
		
		'cancello il singolo file xml elaborato per evitare che venga recuperato nel poll successivo
		call fs.DeleteFile(x.Path,true)
	next
end if

'all'interno del ciclo recupero i dati dei singoli sms da inviare e li inoltro al gestore sms
sData = ""
For Each x In objDictSms
	sData = sData & x.toString
	response.write(x.toString&"<br>")
Next

url = "http://www.sitofornitoresms.com/pagina.php"

Set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
objHttp.open "POST", url, true
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.send sData
response = objHttp.ResponseText
Set objHttp = nothing

'se la response inviata dal fornitore è parlante, la invio via mail all'amministratore del sito per statistiche sugli invii effettuati


Set objXML = nothing
set fo=nothing
set fs=nothing
%>