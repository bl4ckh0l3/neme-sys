<html><title>Utenti online</title><body>
<p>
Who is Online
</p>

    <%
    on error resume next
    response.write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" width=""100%"">"        
    response.write "<tr><td></td><td><strong>CHI</strong></td><td><strong>QUANDO</strong></td><td><strong>DOVE</strong></td></tr>"
    aSessions = dOnlineUsers.Keys
    for iUser = 0 to dOnlineUsers.Count - 1
        sKey = aSessions(iUser)
        sUserInfo = dOnlineUsers.Item(sKey)
        aUserInfo = split(sUserInfo, "<|>")
        
        sUserName = aUserInfo(0)
        sLastActionTime = aUserInfo(1)
        sLastPageViewed = aUserInfo(2)
        
        if sUserInfo <> "" then
        iUsrCount = iUsrCount + 1
        response.write "<tr><td align=""right"">" & iUsrCount & ".</td><td> " & sUserName & "</td><td>" & sLastActionTime & "</td>"
        response.write "<td>" & sLastPageViewed & "</td></tr>"
        end if
        
    next
    response.write "</table>"
    %>

<br><br>
</body></html>

