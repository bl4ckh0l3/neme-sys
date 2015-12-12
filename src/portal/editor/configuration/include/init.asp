<%
if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
Set objUserLoggedTmp = nothing

Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if

Dim objConfig
Set objConfig = New ConfigClass

if(not(request("key") = "")) then
	Set objLogger = New LogClass
	call objConfig.updateConfigValue(request("key"), request("value")) 
	Application(request("key"))= request("value")
	call objLogger.write("modificata variabile di sistema--> key: "&request("key")&"; value: "&request("value"), objUserLogged.getUserName(), "info")
	Set objLogger = nothing
	
	Dim objFSO,configFile	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")		
	Set configFile=objFSO.OpenTextFile(Server.MapPath(Application("baseroot")&"/public/conf/nemesi_config.xml"), 1)		
	Dim m_xmld, m_node, newAttr
	'Create the XML Document: Msxml2.DOMDocument.3.0 oppure Microsoft.XMLDOM
	Set m_xmld = Server.CreateObject("Microsoft.XMLDOM")
	'Load the Xml file
	m_xmld.loadXML(configFile.readAll())	
	configFile.Close
	'Get the servername of the xml file
	Set m_node = m_xmld.SelectSingleNode("/config/"&request("key"))	
	call m_node.setAttribute("attr_"&request("key"),request("value"))	
	m_xmld.save(Server.MapPath(Application("baseroot")&"/public/conf/nemesi_config.xml"))	
	
	Set m_node = nothing
	Set m_xmld = nothing	
	Set configFile=Nothing			
	Set objFSO = nothing
end if
Set objUserLogged = nothing
%>