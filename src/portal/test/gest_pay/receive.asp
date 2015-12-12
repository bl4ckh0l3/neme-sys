<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
</head>
<body>
<%
'INIZIO SCRIPT PER DECRITTOGRAFIA
'DA NON MODIFICARE


'VENGONO LETTI I PARAMETRI IN INPUT E VIENE DECRIPTATO IL
'PARAMETRO B

parametro_a = trim(request("a"))
parametro_b = trim(request("b"))


'PARTE DA NON MODIFICARE
Dim objdeCrypt
'Sintassi Oggetto Java
'Set objdeCrypt = GetObject("java:GestPayCrypt")
'Sintassi Oggetto COM
Set objdeCrypt =Server.Createobject("GestPayCrypt.GestPayCrypt")
'Sintassi Oggetto COM High Security
'Set objdeCrypt =Server.Createobject("GestPayCryptHS.GestPayCryptHS")


if Err.number <> 0 then 
Response.Write Err.number & Err.description
end if

objdeCrypt.SetShopLogin(parametro_a)
objdeCrypt.SetEncryptedString(parametro_b)

call objdeCrypt.Decrypt

'DI SEGUITO SI HANNO UNA SERIE DI VARIABILI VALORIZZATE CON I
'DATI RICEVUTI DA GESTPAY DA UTILIZZARE PER L'INTEGRAZIONE CON
'IL PROPRIO SISTEMA


myshoplogin=trim(objdeCrypt.GetShopLogin)
mycurrency=objdeCrypt.GetCurrency
myamount=objdeCrypt.GetAmount
myshoptransactionID=trim(objdeCrypt.GetShopTransactionID)
mybuyername=trim(objdeCrypt.GetBuyerName)
mybuyeremail=trim(objdeCrypt.GetBuyerEmail)
mytransactionresult=trim(objdeCrypt.GetTransactionResult)
myauthorizationcode=trim(objdeCrypt.GetAuthorizationCode)
myerrorcode=trim(objdeCrypt.GetErrorCode)
myerrordescription=trim(objdeCrypt.GetErrorDescription)
myerrorbanktransactionid=trim(objdeCrypt.GetBankTransactionID)
myalertcode=trim(objdeCrypt.GetAlertCode)
myalertdescription=trim(objdeCrypt.GetAlertDescription)
mycustominfo=trim(objdeCrypt.GetCustomInfo)

response.Write(myshoplogin&"<br>")
response.Write(mycurrency&"<br>")
response.Write(myamount&"<br>")
response.Write(myshoptransactionID&"<br>")
response.Write(mybuyername&"<br>")
response.Write(mytransactionresult&"<br>")
response.Write(myauthorizationcode&"<br>")
response.Write(myerrorcode&"<br>")
response.Write(myerrordescription&"<br>")
response.Write(myerrorbanktransactionid&"<br>")
response.Write(myalertcode&"<br>")
response.Write(myalertdescription&"<br>")
response.Write(mycustominfo&"<br>")

'FINE SCRIPT DI DECRITTOGRAFIA

%>


</body>
</html>