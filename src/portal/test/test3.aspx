<%@ Page Language="C#" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title></title>
</head>
<body>

<%
GestPayCrypt.GestPayCrypt gpc = new GestPayCrypt.GestPayCrypt();
string shopLogin = "GESPAY47944";
string currency = "242";
string amount = "1256.28";
string shopTransactionId = "34az85ord19";
string buyerName = "Denis Testa";
//string EncryptedString = "xxxxxxxxx";
 
gpc.SetShopLogin(shopLogin);
gpc.SetCurrency(currency);
gpc.SetAmount(amount);
gpc.SetShopTransactionID(shopTransactionId);
gpc.SetBuyerName(buyerName);
//gpc.SetEncryptedString(EncryptedString);


gpc.Encrypt();
//string ErrorDesc = gpc.getErrorDescription();
//if (gpc.getErrorCode().Equals("0"))
//{
//string a = gpc.getShopLogin();
string b = gpc.getEncryptedstring();
//}
//else
//{

//TESTRESULT.Text = "ErrorCode:" + gpc.getErrorCode() + "<br />" + " ErrorDesc:" + ErrorDesc;
//}

%>

</body>
</html>