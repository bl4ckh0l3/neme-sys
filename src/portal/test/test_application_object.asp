<%
'response.write("typename(Application(languageResources)): "& typename(languageResources)&"; size: "&languageResources.count&"<br><br>")

'response.write("typename(Application(Session(langElements))): "& typename(Session("langElements"))&"; size: "&Session("langElements").count&"<br><br>")

'for each x in Session("langElements")
'	languageResources.remove(x)
'	languageResources.add x, Session("langElements")(x)
'next

'languageResources.add "IT", Session("langElements")

'languageResources.removeAll
'languageResources.add "IT", Session("langElements")

'response.write("portal.commons.errors.label.error: "&languageResources.item("IT").item("portal.commons.errors.label.error")&"<br><br>")
'languageResources.item("IT").remove("portal.commons.errors.label.error")
'languageResources.item("IT").add "portal.commons.errors.label.error", "CAMBIATO ERROR"

'response.write("portal.commons.errors.label.error: "&languageResources.item("IT").item("portal.commons.errors.label.error")&"<br><br>")

'response.write("typename(Application(languageResources)): "& typename(languageResources)&"; size: "&languageResources.count&"<br><br>")

'response.write("refresh_currency_time: "&Application("refresh_currency_time")&"<br>")



'response.write("Server.MapPath: "&Server.MapPath(Application("dir_down_prod"))&"<br>")


'dim fs,path
'set fs=Server.CreateObject("Scripting.FileSystemObject")
'path=fs.GetAbsolutePathName(Application("dir_down_prod"))
'response.write("www.blackholenet.com"&Application("dir_down_prod")&"<br>")
'response.write(path)


'response.write("APPL_MD_PATH: "&Request.ServerVariables("APPL_MD_PATH")&"<br>")
'response.write("INSTANCE_META_PATH: "&Request.ServerVariables("INSTANCE_META_PATH")&"<br>")


'Function ToRootedVirtual(relativePath)
 '   Dim applicationMetaPath : applicationMetaPath = Request.ServerVariables("APPL_MD_PATH")
  '  Dim instanceMetaPath : instanceMetaPath = Request.ServerVariables("INSTANCE_META_PATH")
   ' Dim rootPath : rootPath = Mid(applicationMetaPath, Len(instanceMetaPath) + Len("/ROOT/"))
    'ToRootedVirtual = rootPath + relativePath
'End Function

'Function ToAppRelative(virtualPath)
 '       Dim sAppMetaPath : sAppMetaPath = Request.ServerVariables("APPL_MD_PATH")
  '      Dim sInstanceMetaPath: sInstanceMetaPath = Request.ServerVariables("INSTANCE_META_PATH")
   '     ToAppRelative = "~/" & Mid(virtualPath, Len(sAppMetaPath) - Len(sInstanceMetaPath) - 3)
'End Function

'response.write(ToRootedVirtual("http://www.blackholenet.com"&Application("dir_down_prod"))&"<br>")
'response.write(ToAppRelative(ToRootedVirtual("http://www.blackholenet.com"&Application("dir_down_prod")))&"<br>")


'Set objFSO = Server.CreateObject("Scripting.FileSystemObject")		
'Set configFile=objFSO.OpenTextFile(Server.MapPath("/public/conf/nemesi_config.xml"), 1)	

'Dim m_xmld, m_node, dbconn

'Create the XML Document: Msxml2.DOMDocument.3.0 oppure Microsoft.XMLDOM
Set m_xmld = Server.CreateObject("Microsoft.XMLDOM")
m_xmld.async = False

'Load the Xml file
'm_xmld.loadXML(configFile.readAll())
m_xmld.Load "http://www.blackholenet.com/public/conf/nemesi_config.xml"

response.write("URL: "&m_xmld.URL&"<br>")
response.write("BaseName: "&m_xmld.BaseName&"<br>")

'Get the servername of the xml file
'Set m_node = m_xmld.SelectSingleNode("/config/srt_dbconn")

'dbconn = m_node.getAttribute("attr_srt_dbconn")
'Application("srt_dbconn") = dbconn

'configFile.Close

'Set m_node = nothing
Set m_xmld = nothing	
'Set configFile=Nothing			
'Set objFSO = nothing


response.write("Server.MapPath: "&Server.MapPath("www.blackholenet.com/public/conf/nemesi_config.xml")&"<br>")
%>