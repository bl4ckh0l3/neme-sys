<html>
<head>
<title>PHP Test</title>
</head>
<body>

<form name="login" action="http://www.blackholenet.com/public/catalog/admin/login.php?action=process" method="post">
<table border="0" width="100%" cellspacing="0" cellpadding="2">
  <tr>
    <td class="infoBoxContent">Username:<br><input type="text" name="username" value="administrator"></td>
  </tr>
  <tr>
    <td class="infoBoxContent"><br>Password:<br><input type="password" name="password" maxlength="40" value="admin"></td>
  </tr>
  <tr>
    <td align="center" class="infoBoxContent"><br><input type="submit" value="Login" /></td>
  </tr>
</table>
</form>

<?php 
$GLOBAL["BASEROOT"] = "/";


$_ENV["PROVA"] = "miao";

echo "DOCUMENT_ROOT: ".$_SERVER["DOCUMENT_ROOT"]."<br/>";

echo "SERVER_NAME: ".$_SERVER["SERVER_NAME"]."<br/>";

echo "BASEROOT: ".$GLOBAL["BASEROOT"]."<br/>";

echo "PROVA: ".$_ENV["PROVA"]."<br/>";
?>

<a href="test2.php">GO</a>

<?php //phpinfo(); ?>


</body>
</html>