<%
Response.Expires = -1
Server.ScriptTimeout = 1200

Set Upload = Server.CreateObject("Persits.Upload")

' Do not throw the "Wrong ContentType error first time out
Upload.IgnoreNoPost = True

Count = Upload.Save(Server.MapPath("/mdb-database/"))

If Count > 0 Then
	Response.Write Count & " file(s) caricati."
End If
%>

<HTML><title>AspUpload: Upload diretto con pagina unica</title> 
<BODY BGCOLOR="#FFFFFF">

<h3>AspUpload: Upload diretto con pagina unica</h3>

	<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="form-upload-unico.asp"> 
		<INPUT TYPE="FILE" SIZE="40" NAME="FILE1"><BR> 
		<INPUT TYPE="FILE" SIZE="40" NAME="FILE2"><BR> 
		<INPUT TYPE="FILE" SIZE="40" NAME="FILE3"><BR> 
	<INPUT TYPE=SUBMIT VALUE="Upload!">
	</FORM>
</BODY> 
</HTML>