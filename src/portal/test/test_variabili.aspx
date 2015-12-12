<%@ Page AspCompat="true" Language="VB" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
</head>
<body>
<%

response.write(String.Concat("<br>Application(servername): ",Application("servername") ,"<br>"))
response.write(String.Concat("<br>Application(uri): ",Application("uri") ,"<br>"))
response.write(String.Concat("<br>Application(TestCounterCurrency): ",Application("TestCounterCurrency") ,"<br>"))
response.write(String.Concat("<br>Application(TestCounterDemo): ",Application("TestCounterDemo") ,"<br>"))
response.write(String.Concat("<br>Application(filename): ",Application("filename") ,"<br>"))
response.write(String.Concat("<br>Application(currencysrvname): ",Application("currencysrvname"),"<br>"))
response.write(String.Concat("<br>Application(uriString): ",Application("uriString") ,"<br>"))
response.write(String.Concat("<br>Application(uriString2): ",Application("uriString2") ,"<br>"))
response.write(String.Concat("<br>Application(utils): ",Application("utils") ,"<br>"))
response.write(String.Concat("<br>Application(error): ",Application("error") ,"<br>"))
response.write(String.Concat("<br>Application(uriString): ",Application("uriString") ,"<br>"))

%>

</body>
</html>
