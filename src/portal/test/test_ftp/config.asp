<%
'///////////////////////////////////////////////////////////////////////////////////
'connection string

Set MyConn=Server.CreateObject("ADODB.Connection")
'MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("/mdb-database/login.mdb")
MyConn.Open "driver={MySQL ODBC 3.51 Driver};uid=Sql198279;pwd=a34d7876;database=Sql198279_1;Server=62.149.150.77;port=3306"

'///////////////////////////////////////////////////////////////////////////////////


'///////////////////////////////////////////////////////////////////////////////////
'cleanup routines

Sub CleanUp(RS)
  RS.Close
  MyConn.Close
  Set RS = Nothing
  Set MyConn = Nothing
End Sub

Sub CleanUp2()
  MyConn.Close
  Set MyConn = Nothing
End Sub

'////////////////////////////////////////////////////////////////////////////////////
%>
