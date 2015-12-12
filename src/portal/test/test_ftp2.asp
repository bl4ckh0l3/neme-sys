<%
Dim strHost, strUser, strPass, strMode, LocalDir, RemoteDir
Dim Output, ReturnCode, strScript
Dim COMMAND_FTP
COMMAND_FTP = "ftp.exe -i -s:"
strHost = "ftp.blackholenet.com"
strUser = "1898894@aruba.it"
strPass = "y46cczq8"
strMode = "ascii" '=== "ascii" / "binary" 
'LocalDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & "published"
LocalDir = ""
RemoteDir = "/public/"

'=====================
'Function FTP( strCMD )
Function FTP()
'=====================
'=== Build a command script, FTPs with it, deletes it
Dim objFSO, strFile, objTempFldr, objFile, objRegExp
Dim objShell, WSX, ReturnCode, Output, strLog, strErrorLog
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
  
set objTempFldr = objFSO.GetSpecialFolder(2) 
strFile = objFSO.GetTempName

strFile = objTempFldr & "\" & strFile & ".ftp"
response.Write("strFile: "&strFile&"<br>")
if not objFSO.FileExists( strFile ) then objFSO.CreateTextFile( strFile )
Set objFile = objFSO.OpenTextFile( strFile, 2, True )

objFile.WriteLine( strUser )
objFile.WriteLine( strPass )
If LocalDir <> "" Then objFile.WriteLine( "lcd " & LocalDir )
If RemoteDir <> "" Then objFile.WriteLine( "cd " & RemoteDir )
'objFile.WriteLine( Mode )

'objFile.WriteLine( strCMD )

'** esperimento
objFile.WriteLine "prompt"
objFile.WriteLine "put " & "paypal_cert_pem.txt"
objFile.WriteLine( "bye" )
objFile.Close()
Set objShell = Server.CreateObject("WScript.Shell")

set WSX = objShell.Exec( COMMAND_FTP & strFile & " " & strHost )
set ReturnCode = WSX.StdErr
set Output = WSX.stdOut
strErrorLog = objTempFldr.Path & "ftpErrors.txt"
strLog = objTempFldr.Path & "ftpLog.txt"
 
Set objFile = objFSO.OpenTextFile( strErrorLog, 2, True )
objFile.Write( ReturnCode.ReadAll() )
objFile.Close()
 
Set objFile = objFSO.OpenTextFile( strLog, 2, True )
objFile.Write( Output.ReadAll() )
objFile.Close()
set objFSO = nothing
set objFile = nothing
 
'objFSO.DeleteFile strFile, True
set objFSO = nothing

Set objRegExp = New RegExp   
objRegExp.IgnoreCase = True 
objRegExp.Pattern = "not connected|invalid command|error"
 
If (objRegExp.Test( Output.ReadAll ) = True ) or (objRegExp.Test( ReturnCode.ReadAll ) ) Then
   FTP = False
Else
   FTP = True
End If

Set objRegExp = nothing
End Function

result = FTP()
response.Write("result: "&result)

'strCommands = “MKD JUSTINS_BABY_PICTURES” & vbCrLf & _
'“MKD LOOKED_LIKE_MONKEY” & vbCrLf & _
'“PUT simian1.jpg” & vbCrLf & _
'“PUT banana.gif”

%>