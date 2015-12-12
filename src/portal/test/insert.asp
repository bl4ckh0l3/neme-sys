<% ' Insert.asp %>
<!--#include virtual="/common/include/objects/ImageUploadClass.asp"-->
<%
  Response.Buffer = True

  ' load object
  Dim load
    Set load = new ImageUploadClass
    
    ' calling initialize method
    load.initialize
    
  ' File binary data
  Dim fileData
    fileData = load.getFileData("file")
  ' File name
  Dim fileName
    fileName = LCase(load.getFileName("file"))
  ' File path
  Dim filePath
    filePath = load.getFilePath("file")
  ' File path complete
  Dim filePathComplete
    filePathComplete = load.getFilePathComplete("file")
  ' File size
  Dim fileSize
    fileSize = load.getFileSize("file")
  ' File size translated
  Dim fileSizeTranslated
    fileSizeTranslated = load.getFileSizeTranslated("file")
  ' Content Type
  Dim contentType
    contentType = load.getContentType("file")
  ' No. of Form elements
  Dim countElements
    countElements = load.Count
  ' Value of text input field "fname"
  Dim fnameInput
    fnameInput = load.getValue("fname")
  ' Value of text input field "lname"
  Dim lnameInput
    lnameInput = load.getValue("lname")
  ' Value of text input field "profession"
  Dim profession
    profession = load.getValue("profession")  
    
  ' destroying load object
  Set load = Nothing
%>

<html>
<head>
  <title>Inserts Images into Database</title>
  <style>
    body, input, td { font-family:verdana,arial; font-size:10pt; }
  </style>
</head>
<body>
  <p align="center">
    <b>Inserting Binary Data into Database</b><br>
    <a href="insert.htm">insert again data click here</a>
  </p>
  
  <table width="700" border="1" align="center">
  <tr>
    <td>File Name</td><td><%= fileName %></td>
  </tr><tr>
    <td>File Path</td><td><%= filePath %></td>
  </tr><tr>
    <td>File Path Complete</td><td><%= filePathComplete %></td>
  </tr><tr>
    <td>File Size</td><td><%= fileSize %></td>
  </tr><tr>
    <td>File Size Translated</td><td><%= fileSizeTranslated %></td>
  </tr><tr>
    <td>Content Type</td><td><%= contentType %></td>
  </tr><tr>
    <td>No. of Form Elements</td><td><%= countElements %></td>
  </tr><tr>
    <td>First Name</td><td><%= fnameInput %></td>
  </tr><tr>
    <td>Last Name</td><td><%= lnameInput %></td>
  </tr>
  <tr>
    <td>Profession</td><td><%= profession %></td>
  </tr>
  </table><br><br>
  
  <p style="padding-left:220;">
  <%= fileName %> data received ...<br>
  <%
  Const adCmdText = 1
  Const adOpenDynamic = 2
  Const adLockOptimistic = 3
  Const adOpenKeyset = 1

    ' Checking to make sure if file was uploaded
    If fileSize > 0 Then
    
      ' Connection string
      Dim connStr	
	connStr = "driver={MySQL ODBC 3.51 Driver};uid=Sql198279;pwd=a34d7876;database=Sql198279_1;port=3306;Server=62.149.150.77"
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = connStr
	objConn.open()	
	
      ' Recordset object
      Dim rs
        Set rs = Server.CreateObject("ADODB.Recordset")        
	rs.Open "utenti_images",objConn, 2, 2, 2

        ' Adding data
        rs.AddNew
          rs("id_utente") = 55
          rs("filename") = fileName
          rs("file_size") = fileSize
          rs("file_data").AppendChunk fileData
          'rs("file_data") = fileData
          rs("content_type") = contentType
          'rs("First Name") = fnameInput
          'rs("Last Name") = lnameInput
          'rs("Profession") = profession
        rs.Update        
        rs.Close
        Set rs = Nothing
	Set objConn = nothing
        
      Response.Write "<font color=""green"">File was successfully uploaded..."
      Response.Write "</font>"
    Else
      Response.Write "<font color=""brown"">No file was selected for uploading"
      Response.Write "...</font>"
    End If
      
      
    If Err.number <> 0 Then
      Response.Write "<br><font color=""red"">Something went wrong..."
      Response.Write "</font>"
    End If
  %>
  </p>
  
  <br>
  <table border="0" align="center">
  <tr>
  <form method="POST" enctype="multipart/form-data" action="Insert.asp">
  <td>First Name :</td><td>
    <input type="text" name="fname" size="40" ></td>
  </tr>
  <td>Last Name :</td><td>
    <input type="text" name="lname" size="40" ></td>
  </tr>
  <td>Profession :</td><td>
    <input type="text" name="profession" size="40" ></td>
  </tr>
  <td>File :</td><td>
    <input type="file" name="file" size="40"></td>
  </tr>
  <td> </td><td>
    <input type="submit" value="Submit"></td>
  </tr>
  </form>
  </tr>
  </table>

</body>
</html>