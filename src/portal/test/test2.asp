<%
on error resume next
myDomain = "localhost"
myrewritedir = "code"
 
Set objW3SVC = GetObject("IIS://localhost/W3SVC")
For Each objSITE in objW3SVC
  If objSITE.class = "IIsWebServer" Then
 
      websiteNameArr = objSITE.ServerBindings
        for j = 0 to Ubound(websiteNameArr)
          websiteName = websiteNameArr(j)
 
           If instr(websiteName,myDomain) > 0 then
 
            Set objIISNewDir = GetObject("IIS://localhost/W3SVC/" & objSite.Name & "/root")
            Set CodeDir = objIISNewDir.Create("IIsWebDirectory",myrewritedir )
            CodeDir.SetInfo
            Set objIISNewDir = Nothing
 
            Set objIISRewriteRootDir = GetObject("IIS://localhost/W3SVC/" & objSite.Name & "/root/" & myrewritedir)
              CustomErrors = objIISRewriteRootDir.HttpErrors
              For i = 0 To UBound(CustomErrors)
                  If Left(CustomErrors(i),3) = "404" then
                      CustomErrors(i) = "404,*,URL,/" & myrewritedir & "/rewrite.asp"
                      objIISRewriteRootDir.HttpErrors = CustomErrors
                	  objIISRewriteRootDir.SetInfo
                      Exit For
                   End If
              Next
            Set objIISRewriteRootDir = Nothing
          End if
        Next
 
  End if
Next
%>
