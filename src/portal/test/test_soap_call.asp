<!-- include virtual="/common/include/i_soapcall.asp" -->
<!--
<HTML>
<HEAD>
</HEAD>
<BODY>
<%
'response.write("ASPNETResourceList: "&Application("ASPNETResourceList"))

'Dim     ASPNETResources
'If len( Application("ASPNETResourceList") )>0 then    'we have our
latest resources

'    REM -- check to see if they expired
'    If DateDiff("h",Now(),Application("ASPNETResourcesUpdated")) >
Application("ASPNETExpires") Then
'        REM -- we need to update the latest resurces
'        ASPNETResources = GetASPNetResources()
'        Application.Lock
'        Application("ASPNETResourcesUpdated")=Now()
'        Application("ASPNETResourceList")=ASPNETResources
'        Application.UnLock
'    End if 'datediff...
'
'Else    'for some reason the application level variable is empty, fill it.
'    ASPNETResources = GetASPNetResources()
'    Application.Lock
'    Application("ASPNETResourcesUpdated")=Now()
'    Application("ASPNETResourceList")=ASPNETResources
'    Application.UnLock

'End if 'len(..

'Response.Write     Application("ASPNETResourceList")

%>
<P>&nbsp;</P>

</BODY>
</HTML>