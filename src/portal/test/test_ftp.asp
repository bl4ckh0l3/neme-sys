<%@ Language=VBScript %>
<%
' FTP via ASP without using 3rd-party components
' Ben Meghreblian 15th Jan 2002
' benmeg at benmeg dot com / http://benmeg.com/code/asp/ftp.asp.html
'
' This script assumes the file to be FTP'ed is in the same directory as this script.
' It should be obvious how to change this (*hint* change the lcd line).
' You may specify a wildcard in ftp_files_to_put (e.g. *.txt).

' NB: You need to have C:\winnt\system32\wshom.ocx registered to use the WSCRIPT.SHELL object.
' It is registered by default, but is sometimes removed for security reasons (no kidding!).
' You will also need cmd.exe in the path, which again is there, unless the box is locked down.
' Check with your web host/resident sysadmin if in doubt.
'
' NB: This script was originally written in response to a thread on a Wrox ASP mailing list.
' At the time, I was hosting on a shared NT4/IIS4 box and the script worked fine. Since I wrote
' it, several people have got in contact asking why it doesn't work on later versions of either
' Windows or IIS. The answer is probably either as mentioned in the above NB, or to do with
' firewalls restricting outbound traffic from and/or to certain ports. This said, many people
' have successfully used this code to FTP to/from Windows 2000/Windows XP boxes running IIS5/IIS6.
Dim objFSO, objTextFile, oScript, oScriptNet, oFileSys, oFile, strCMD, strTempFile, strCommandResult
Dim ftp_address, ftp_username, ftp_password, ftp_physical_path, ftp_files_to_put

' Edit these variables to match your specifications
ftp_address          = "ftp.blackholenet.com"
ftp_username         = "1898894@aruba.it"
ftp_password         = "y46cczq8"
ftp_remote_directory = "/blackholenet.com/public/" ' Leave blank if uploading to root directory
ftp_files_to_put     = "paypal_cert_pem.txt"     ' You can use wildcards here (e.g. *.txt)
On Error Resume Next
Set oScript = Server.CreateObject("WSCRIPT.SHELL")
Set oFileSys = Server.CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")

response.Write("oScript: "&typename(oScript)&"<br>")
response.Write("oFileSys: "&typename(oFileSys)&"<br>")
response.Write("objFSO: "&typename(objFSO)&"<br>")

' Build our ftp-commands file
Set objTextFile = objFSO.CreateTextFile(Server.MapPath("test.ftp"))
response.Write("Server.MapPath: "&Server.MapPath("test.ftp")&"<br>")
objTextFile.WriteLine "lcd " & Server.MapPath(".")
'objTextFile.WriteLine "open " & ftp_address
objTextFile.WriteLine ftp_username
objTextFile.WriteLine ftp_password

' Check to see if we need to issue a 'cd' command
If ftp_remote_directory <> "" Then
   objTextFile.WriteLine "cd " & ftp_remote_directory
End If

objTextFile.WriteLine "prompt"

' If the file(s) is/are binary (i.e. .jpg, .mdb, etc..), uncomment the following line
'objTextFile.WriteLine "binary"

' If there are multiple files to put, we need to use the command 'mput', instead of 'put'
If Instr(1, ftp_files_to_put, "*",1) Then
   objTextFile.WriteLine "mput " & ftp_files_to_put
Else
   objTextFile.WriteLine "put " & ftp_files_to_put
End If
objTextFile.WriteLine "bye"
objTextFile.Close
Set objTextFile = Nothing
' Use cmd.exe to run ftp.exe, parsing our newly created command file
'strCMD = "ftp.exe -s:" & Server.MapPath("test.ftp")
strCMD = "ftp.exe -i -s:" & Server.MapPath("test.ftp")
set objTempFldr = objFSO.GetSpecialFolder(2)

response.Write("objTempFldr: "&objTempFldr&"<br>")

strTempFile = objTempFldr & "" & oFileSys.GetTempName( )
' Pipe output from cmd.exe to a temporary file (Not :| Steve)
'Call oScript.Run ("cmd.exe /c " & strCMD & " > " & strTempFile, 0, True)
set WSX = oScript.Exec(strCMD & " " & ftp_address )
response.Write("WSX: "&typename(WSX)&"<br>")

set ReturnCode = WSX.StdErr
set Output = WSX.stdOut
strErrorLog = Server.MapPath("ftpErrors.txt")
strLog = Server.MapPath("ftpLog.txt")
 
Set objFile = objFSO.OpenTextFile( strErrorLog, 2, True )
objFile.Write( ReturnCode.ReadAll() )
objFile.Close()
 
Set objFile = objFSO.OpenTextFile( strLog, 2, True )
objFile.Write( Output.ReadAll() )
objFile.Close()
set objFile = nothing

'Set oFile = oFileSys.OpenTextFile (strTempFile, 1, False, 0)

On Error Resume Next
' Grab output from temporary file
'strCommandResult = Server.HTMLEncode( oFile.ReadAll )
'oFile.Close
' Delete the temporary & ftp-command files
'Call oFileSys.DeleteFile( strTempFile, True )
'Call objFSO.DeleteFile( Server.MapPath("test.ftp"), True )
Set oFileSys = Nothing
Set objFSO = Nothing
' Print result of FTP session to screen
response.write("end script")
response.write("strCommandResult: "&strCommandResult)
Response.Write( Replace( strCommandResult, vbCrLf, "<br>", 1, -1, 1) )
%>
