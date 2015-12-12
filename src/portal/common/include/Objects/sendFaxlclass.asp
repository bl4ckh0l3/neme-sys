<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Type Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->

<%
Class SendFaxClass
	Sub invia(ByVal numero, ByVal file, ByVal rascl, ByVal utente)
		Dim FaxServer As Object
		Dim FaxDoc As Object

		FaxServer = Server.CreateObject("Faxserver.FaxServer")
		'*** sostituire il nomer del server corrente
		FaxServer.Connect("localhost")
		FaxDoc = FaxServer.CreateDocument(file)
		Dim FAXCL = ""
		Dim i
		'pulisco il numero di fax da caratteri non numerici
		For i = 1 To Len(numero)
			If Asc(Mid(numero, i, 1)) >= 48 And Asc(Mid(numero, i, 1)) <= 57 Then
				FAXCL = FAXCL & Mid(numero, i, 1)
			End If
		Next
		
		'aggiungo 0 davanti al numero per uscire dal mio centralino
		FaxDoc.FaxNumber = "0" & FAXCL
		'in recipient name ci metto delle informazioni che mi servono
		FaxDoc.RecipientName = Right(rascl, 6) & "#" & file & "#" & utente
		
		Dim JobID
		On Error Resume Next
		JobID = FaxDoc.Send()

		If Err.Number <> 0 Then
			Response.Write("Impossibile inviare il fax (" & Err.Description & ")")
		Else
			Response.Write("Fax inviato con successo - JobID = " & JobID)
			Response.Write(" Cliente: " & Right(rascl, 6) & " Fax: " & FAXCL)
		End If

		FaxDoc = Nothing
		FaxServer.Disconnect()
		FaxServer = Nothing
	End Sub
End Class


'***************************** TEST DI ESEMPIO *****************************
'*** chiamo la funzione di invio fax
'Call invia(numero, file, rascl, utente)

'*** uccido acrobat
'Dim proc As System.Diagnostics.Process
'Dim pList() As System.Diagnostics.Process
'pList = System.Diagnostics.Process.GetProcessesByName("AcroRd32")
'For Each proc In pList
'   proc.Kill()
'Next
'***************************** FINE TEST DI ESEMPIO *****************************
%>