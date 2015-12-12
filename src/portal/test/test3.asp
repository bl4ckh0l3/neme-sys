<%'On Error Resume Next%>
<!--#include virtual="/common/include/IncludeShopObjectList.inc" -->
<!--include file ="common/include/Objects/objPageCache.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
</head>
<body>
<%
'TENTATIVO PER USARe I SOTTODOMINI PEr LA GESTIONE DELLE LINGUE
'OTTIMIZZATA PER I MOTORI DI RICERCA
'Dim http_host, subdomainlang
'http_host = Request.ServerVariables("HTTP_HOST")
'subdomainlang = Left(http_host,InStr(0,http_host,".",1)-1) 
'subdomainlang = Ucase(subdomainlang)
'response.Write(subdomainlang)
'Server.Transfer(pageRedirect & "?lang="&subdomainlang)
%> 

<form name="test" method="post" action="">
<a href="javascript:document.test.submit();">vai</a>
</form>

<%
'Const FOLDER_PATH=Server.MapPath(Application("dir_down_prod"))
FOLDER_PATH=Server.MapPath(Application("dir_down_prod"))
response.write(Application("dir_down_prod"))
response.write("<br><br>"&FOLDER_PATH)

'response.Write("<br><br>"&Request.ServerVariables("SERVER_NAME"))


'response.write("<br><br>"&Server.MapPath("www.blackholenet.com"&Application("dir_down_prod")))


'Set objFSO = Server.CreateObject("Scripting.FileSystemObject")


'response.write("<br><br>URL: www.blackholenet.com"&Application("dir_upload_user"))

'response.write("<br><br>exists: ")
'response.write(objFSO.FolderExists(Application("dir_upload_user")))

'Set objFSO = nothing

response.write("<br><br>path to backup: ")
response.write(Server.MapPath("/blackholenet.com_Backup_Giornaliero/common/swf/"))
%>

</body>
</html>