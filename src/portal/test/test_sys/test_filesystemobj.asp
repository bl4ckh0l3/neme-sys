<%
Dim fso
Dim inFile
Dim outFile
Dim riga
Dim fileName
Dim newText


Function ReadFile(sFilePathAndName) 
   dim sFileContents 

   Set oFS = Server.CreateObject("Scripting.FileSystemObject") 

   If oFS.FileExists(sFilePathAndName) = True Then        
      Set oTextStream = oFS.OpenTextFile(sFilePathAndName,1)        
      sFileContents = oTextStream.ReadAll      
      oTextStream.Close 
      Set oTextStream = nothing    
   End if 
   
   Set oFS = nothing 

   ReadFile = sFileContents   
End Function


'nemesiConfigFile = "/public/conf/nemesi_config.xml"
'fileName = Server.MapPath(nemesiConfigFile)
'response.write(fileName&"<br/>")

' NOTA:
' si suppone che 'fileName' contenga il nome del 
' file da leggere.

' creo il FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

'************* CREO UN NUOVO FILE nemesi_config.xml dove inserisco il nome del server da recuperare quando necessario


'Set configFile=fso.OpenTextFile(fileName, 2, True)
	'configFile.writeLine("<config>")
		'configFile.writeLine("<currencysrvname>"&Application("srt_default_server_name")&"</currencysrvname>")
		'configFile.writeLine("<strdbconn>"&Application("srt_dbconn")&"</strdbconn>")
	'configFile.writeLine("</config>")

'configFile.Close
'Set configFile=Nothing

'response.write(ReadFile(fileName))

'response.write("Application(srt_default_server_name): " & Application("srt_default_server_name")&"<br>")
'response.write("Application(srt_dbconn): " & Application("srt_dbconn")&"<br>")
'response.write("Application(test): " & Application("test")&"<br>")


'*** recupero la lista di tutte le sottocartelle della direcotry dei template

folderspec = server.mappath(Application("baseroot")&Application("dir_upload_templ"))
lang_code=Ucase("sv")
response.write(lang_code & "<br>")

Set fold = fso.GetFolder(folderspec) 
for each subfolder in fold.subFolders
	subfold=subfolder.Name
	subfoldpath=fso.BuildPath(folderspec,subfold)
	Response.Write("subfold: "&subfold & "<br>") 
	Response.Write("subfoldpath: "&subfoldpath & "<br>") 
	
	subfoldpathnew=fso.BuildPath(subfoldpath,lang_code)
	Response.Write("subfoldpathnew: "&subfoldpathnew & "<br>") 
	Response.Write("subfoldpathnew exist: "&fso.FolderExists(subfoldpathnew) & "<br>") 
	
	if not(fso.FolderExists(subfoldpathnew)) then
		fso.CreateFolder(subfoldpathnew)
	end if
	Response.Write("subfoldpathnew exist: "&fso.FolderExists(subfoldpathnew) & "<br>") 
	
	'if(fso.FolderExists(subfoldpath&"\include")) then
	'	if not(fso.FolderExists(subfoldpathnew&"\include")) then
	'		fso.CreateFolder(subfoldpathnew&"\include")
	'	end if
	'end if
	
	'** copio i file dalla dir principale alla nuova sottodir della lingua specificata
	fso.CopyFile subfoldpath&"\*.asp",subfoldpathnew&"\"
	'** copio tutta la dir include nella nuova sottodir della lingua specificata
	fso.CopyFolder subfoldpath&"\include",subfoldpathnew&"\"
	Exit for
next 
set fold = nothing 




' elimino il FileSystemObject
Set fso=Nothing


%>