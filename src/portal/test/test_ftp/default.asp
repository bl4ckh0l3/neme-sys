<%@Language=VBSCript%>
<%Response.Buffer=True%>

<html>

<head>
<title></title>
</head>

<body>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="30%" bgcolor="#C0C0C0">
    <tr>
      <td width="100%" bgcolor="#33CCFF">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%">

<form action="validate.asp" method="Post">
<p align="center"><font face="Andale Mono">
&nbsp;</font><p align="center"><font face="Trebuchet MS">UserName<b> </b></font>
<font face="Andale Mono">
<input type="text" name="username" size="20"><br>
</font><font face="Trebuchet MS">
&nbsp;Password</font><font face="Andale Mono">&nbsp;
<input type="password" name="password" size="20"></font><p align="center">
<font face="Andale Mono">
<br>
<font face="Trebuchet MS">
<input type="submit" value="accedi"></font>
</font>
</form>

        <p><font face="Trebuchet MS" size="2">per accedere è
        necessario eseguire l'accesso</font></td>
    </tr>
    <tr>
      <td width="100%" bgcolor="#33CCFF">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" bgcolor="#808080">
        <p align="center"><font face="Trebuchet MS" size="2">user <b>pippo</b></font></p>
        <p align="center"><font face="Trebuchet MS" size="2">psw&nbsp; <b>pippo</b></font></td>
    </tr>
  </table>
  </center>
</div>

<%
If Session("allow") = False Then
  Response.Write "Non hai eseguito l'accesso."
Else
  Response.Write "accesso eseguito."
  Response.Write "<p><a href=""utility.asp?method=abandon"">Esci</a>"
End If
%>

</body>
</html>
