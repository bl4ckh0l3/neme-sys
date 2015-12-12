  <!--#include file="upload.asp"-->
<%Response.Expires=0
  Response.Buffer = TRUE
  Response.Clear
  byteCount = Request.TotalBytes
  RequestBin = Request.BinaryRead(byteCount)
  Dim Cartella
  'attenzione, unica riga da modificare
  'inserire il percorso della sotto-cartella in public, ESISTENTE, nella quale verranno inseriti i files
  '-----------------------------------------------------
  Cartella = "/app_data/"
  '------------------------------------------------------
  'nota, se vuoi fare upload in cartella nella quale i files siano raggiungibili solo via FTP (per massima sicurezza)
  'puoi cambiare il percorso, ad esempio con "/mdb-database/nomecartella/"
  'fine modifica linkbc 07/07/2008
  
  Dim UploadRequest
  Set UploadRequest = CreateObject("Scripting.Dictionary")
  BuildUploadRequest  RequestBin
  contentType = UploadRequest.Item("blob").Item("ContentType")
  filepathname = UploadRequest.Item("blob").Item("FileName")
  filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
  value = UploadRequest.Item("blob").Item("Value")

  'Create FileSytemObject Component
  Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")


pathEnd = Cartella 
Set MyFile = ScriptObject.CreateTextFile(Server.mappath(pathEnd & filename), true)

 
  For i = 1 to LenB(value)
	 MyFile.Write chr(AscB(MidB(value,i,1)))
  Next
  MyFile.Close
  
  response.write("byteCount: "&byteCount&"<br>")
  response.write("Cartella: "&Cartella&"<br>")
  response.write("contentType: "&contentType&"<br>")
  response.write("filepathname: "&filepathname&"<br>")
  response.write("filename: "&filename&"<br>")
  response.write("value: "&value&"<br>")
  %>

<head>
<title></title>
</head>

<body bgcolor="#FFFFFF">

<p align="center"><font face="Trebuchet MS">
  File "<b><%=filename%></b>" ricevuto con successo </font>
<p align="center"><font face="Verdana" size="2"><a href="inizia.asp">torna</a></font></p>
