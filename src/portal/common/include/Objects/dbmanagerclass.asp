<%
'  classe di utilità per connessione al db

Class DBManagerClass
	Private objConn	

	Private Sub Class_Initialize()
		on error resume next
		Dim strPathConnDB, srtUser, strPwd
			
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.ConnectionString = Application("srt_dbconn")		
		
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			response.Redirect(Application("baseroot")&"/public/layout/include/error.html")
		end if
	End Sub

	Private Sub Class_Terminate()
		on error resume next
		
		If IsObject(objConn) Then
			closeConnection()
			Set objConn = Nothing
		End If
				
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			response.Redirect(Application("baseroot")&"/public/layout/include/error.html")
		end if
	End Sub
	
	Public Sub setConnectionString(strValue)
		objConn.ConnectionString = strValue
	End Sub	
		
	public Property Get openConnection()
		on error resume next		
		objConn.Open()
		Set openConnection = objConn
		
		if Err.number <> 0 then
			'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			response.Redirect(Application("baseroot")&"/public/layout/include/error.html")
		end if				
	End Property
		
	private Sub closeConnection()
		on error resume next
		objConn.Close				
		
		if Err.number <> 0 then
			'response.write "Error closing DB connection: " & Err.description
			response.Redirect(Application("baseroot")&"/public/layout/include/error.html")
		end if
	End Sub
End Class
%>