<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/UTF8Filer.asp" -->
<%
if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
	Set objUserLoggedTmp = nothing

	Dim strRuoloLogged
	strRuoloLogged = objUserLogged.getRuolo()
	if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
	end if
end if	

Dim objLogger, filepath, content, command, fileid
Dim MyUTF8File
Dim objSelLanguage, objLangList
Set objLogger = New LogClass

filepath = request("filepath")
content = request("content")
command = request("command")
fileid = request("fileid")

on Error Resume Next
	'call objLogger.write("BO global file update - filepath: "&filepath & " - content: "&content&" - command: "&command, "system", "debug")

	Select Case command
		Case "loadfile"
			Set MyUTF8File = New UTF8Filer
			MyUTF8File.UnicodeCharset = "UTF-8"
			if not MyUTF8File.LoadFile(filepath) then
				'The filenames are stored here during the open process
				Response.Write("$('#message').html(""<h2>"&MyUTF8File.AbsoluteFileName & " (" & MyUTF8File.VirtualFileName & ")<br>"&MyUTF8File.ErrorText & "<br>"&"</h2>"")")
			else
				MyUTF8File.cTextBuffer2Unicode
				response.write(MyUTF8File.TextBuffer)
			end if
			set MyUTF8File = nothing
		Case "savefile"
			Set MyUTF8File = New UTF8Filer
			MyUTF8File.UnicodeCharset = "UTF-8"
			if not MyUTF8File.LoadFile(filepath) then
				call objLogger.write("BO global file load - Err.description: "&MyUTF8File.AbsoluteFileName & " (" & MyUTF8File.VirtualFileName & ")<br>"&MyUTF8File.ErrorText & "<br>", "system", "error")
			end if
			MyUTF8File.cTextBuffer2Unicode
			MyUTF8File.TextBuffer=content
			'MyUTF8File.cUnicode2UTF8
			if not MyUTF8File.SaveFile(filepath) then
				call objLogger.write("BO global file update - Err.description: "&MyUTF8File.AbsoluteFileName & " (" & MyUTF8File.VirtualFileName & ")<br>"&MyUTF8File.ErrorText & "<br>", "system", "error")
			end if
			set MyUTF8File = nothing
		Case "savefilepart"
			Set objSelLanguage = New LanguageClass
			Set objLangList = objSelLanguage.getListaLanguage()
			Set objSelLanguage = nothing

			Set MyUTF8File = New UTF8Filer
			MyUTF8File.UnicodeCharset = "UTF-8"
			if not MyUTF8File.LoadFile(filepath) then
				call objLogger.write("BO global file load - Err.description: "&MyUTF8File.AbsoluteFileName & " (" & MyUTF8File.VirtualFileName & ")<br>"&MyUTF8File.ErrorText & "<br>", "system", "error")
			end if
			MyUTF8File.cTextBuffer2Unicode
			MyUTF8File.TextBuffer=content
			'MyUTF8File.cUnicode2UTF8
			if not MyUTF8File.SaveFile(filepath) then
				call objLogger.write("BO global file update - Err.description: "&MyUTF8File.AbsoluteFileName & " (" & MyUTF8File.VirtualFileName & ")<br>"&MyUTF8File.ErrorText & "<br>", "system", "error")
			end if

			For Each lang In objLangList
				if(InStrRev(filepath,".inc",-1,1)>0) then
					filepathlang = Mid(filepath,1,InStrRev(filepath,"/include/",-1,1))
					filepathlang = filepathlang & Ucase(objLangList(lang).getLanguageDescrizione()) & "/include/"
					filepathlang = filepathlang & Right(filepath,Len(filepath)-InStrRev(filepath,"/",-1,1))
				else
					filepathlang = Mid(filepath,1,InStrRev(filepath,"/",-1,1))
					filepathlang = filepathlang & Ucase(objLangList(lang).getLanguageDescrizione()) & "/"
					filepathlang = filepathlang & Right(filepath,Len(filepath)-InStrRev(filepath,"/",-1,1))
				end if
				
				if not MyUTF8File.LoadFile(filepathlang) then
					call objLogger.write("BO global file load - Err.description: "&MyUTF8File.AbsoluteFileName & " (" & MyUTF8File.VirtualFileName & ")<br>"&MyUTF8File.ErrorText & "<br>", "system", "error")
				end if
				MyUTF8File.cTextBuffer2Unicode
				MyUTF8File.TextBuffer=content
				'MyUTF8File.cUnicode2UTF8
				if not MyUTF8File.SaveFile(filepathlang) then
					call objLogger.write("BO global file update - Err.description: "&MyUTF8File.AbsoluteFileName & " (" & MyUTF8File.VirtualFileName & ")<br>"&MyUTF8File.ErrorText & "<br>", "system", "error")
				end if
			Next
			
			set MyUTF8File = nothing
			Set objLangList = nothing
		Case "deletefilepart"
			Set objPagePerTemp = new Page4TemplateClass
			Set objSelLanguage = New LanguageClass
			Set objLangList = objSelLanguage.getListaLanguage()
			Set objSelLanguage = nothing

			call objPagePerTemp.deletePagePerTemplate(fileid)

			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			if objFSO.FileExists(Server.MapPath(filepath)) then
				call objFSO.DeleteFile(Server.MapPath(filepath), true)	
			end if

			For Each lang In objLangList
				if(InStrRev(filepath,".inc",-1,1)>0) then
					filepathlang = Mid(filepath,1,InStrRev(filepath,"/include/",-1,1))
					filepathlang = filepathlang & Ucase(objLangList(lang).getLanguageDescrizione()) & "/include/"
					filepathlang = filepathlang & Right(filepath,Len(filepath)-InStrRev(filepath,"/",-1,1))
				else
					filepathlang = Mid(filepath,1,InStrRev(filepath,"/",-1,1))
					filepathlang = filepathlang & Ucase(objLangList(lang).getLanguageDescrizione()) & "/"
					filepathlang = filepathlang & Right(filepath,Len(filepath)-InStrRev(filepath,"/",-1,1))
				end if
				
				if objFSO.FileExists(Server.MapPath(filepathlang)) then
					call objFSO.DeleteFile(Server.MapPath(filepathlang), true)	
				end if
			Next
			
			Set objFSO = nothing
			Set objLangList = nothing
			Set objPagePerTemp =nothing
		Case Else			
	End Select

if(Err.number<>0)then
	call objLogger.write("BO global file update end - Err.description: "&Err.description, "system", "error")
end if

Set objLogger = nothing
%>