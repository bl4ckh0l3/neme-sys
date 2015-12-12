<%
Private Sub Test()
   Dim rs
   Dim n
  For n = 0 To 2
    Set rs = CreateObject("ADODB.Recordset")
    Set cn(n) = CreateObject("ADODB.Connection")
    cn(n).Open Application("srt_dbconn")
    rs.Open "select * from news_find", cn(n)
    rs.Close
    Set rs = Nothing
    cn(n).Close
  Next
End
	
Private Sub Form_Load()
   Dim rs
   Dim n
  For n = 0 To 2
    Set rs = CreateObject("ADODB.Recordset")
    Set cn(n) = CreateObject("ADODB.Connection")
    cn(n).Open Application("srt_dbconn")
    rs.Open "select * from news_find", cn(n)
    rs.Close
    Set rs = Nothing
    cn(n).Close
    Set cn(n) = Nothing
  Next
End Sub

%>