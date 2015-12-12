<!-- #include virtual="/editor/payments/moduli/paypal/PaypalClass.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
</head>
<body>
<p><b>Versione di IIS:</b>
<%Response.Write(Request.ServerVariables("server_software"))%>
</p>

<%
Dim objTest, objName
objName = "SellaClass"
Set objTest= new  objName
response.write(typename(objTest)&"<br>")
Set objTest=nothing

%>

<%
' INIZIO SCRIPT DI CRITTOGRAFIA

'PARTE DA NON MODIFICARE
Dim objCrypt
'Sintassi Oggetto Java
'Set objCrypt = GetObject("java:GestPayCrypt")
'Sintassi Oggetto COM
Set objCrypt =Server.Createobject("GestPayCrypt.GestPayCrypt")
'Sintassi Oggetto COM High Security
'Set objCrypt =Server.Createobject("GestPayCryptHS.GestPayCryptHS")

response.write("typename(objCrypt): "&typename(objCrypt)&"<br>")
response.write("objCrypt: "&objCrypt&"<br>")


if Err.number <> 0 then 
	Response.Write Err.number & Err.description
end if

'PARTE DA MODIFICARE (VALORIZZAZIONE ATTRIBUTI TRANSAZIONE)

'Inserire al posto delle scritte con parentesi quadre [] I dati
'necessari per effettuare la transazione.
'Le righe contenenti i dati contrassegnati come NON OBBLIGATORI
'devono essere eliminate se non utilizzate

'CAMPI OBBLIGATORI

myshoplogin= "GESPAY47944" '"[SHOP LOGIN]" 'Es. 9000001
mycurrency=242 '"[CODICE DIVISA]" 'Es. 242 per euro o 18 lira
myamount="1256.28" '"[IMPORTO SENZA SEPARATORI DI MIGLIAIA CON SEPARATORE PUNTO PER DECIMALI]" 'Es. "1256.28"
myshoptransactionID="34az85ord19" '"[IDENTIFICATIVO TRANSAZIONE]" 'Es. "34az85ord19"

'mycreditcard="1234567812345678" '"[NUMERO CARTA DI CREDITO]"
'myexpirymounth="09" '"[MESE SCADENZA CARTA]"
'myexpiryyear="2010" '"[ANNO SCADENZA CARTA]"


'CAMPI NON OBBLIGATORI (CANCELLARE LE RIGHE NON INTERESSATE)

mybuyername="pippo" '"[NOME E COGNOME ACQUIRENTE]"'Es. "Mario Bianchi"
mybuyeremail="pippo@gmail.com" '"[EMAIL ACQUIRENTE]"'Es. "Mario.bianchi@isp.it"
'mylanguage="[CODICE LINGUA DA UTILIZZARE NELLA COMUNICAZIONE]" 'Es. "3" per spagnolo
'mycustominfo="[PARAMETRI PERSONALIZZATI]" 'Es. "BV_CODCLIENTE=12*P1*BV_SESSIONID=398"




objCrypt.SetShopLogin(myshoplogin)
objCrypt.SetCurrency(mycurrency)
objCrypt.SetAmount(myamount)
objCrypt.SetShopTransactionID(myshoptransactionID)
'objCrypt.SetBuyerName(mybuyername)
'objCrypt.SetBuyerEmail(mybuyeremail)
'objCrypt.SetLanguage(mylanguage)
'objCrypt.SetCustomInfo(mycustominfo)
'objCrypt.SetCardNumber(mycreditcard)
'objCrypt.SetExpMonth(myexpirymounth)
'objCrypt.SetExpYear(myexpiryyear)

response.write("objCrypt.getShopLogin: "&objCrypt.GetShopLogin()&"<br>")
response.write("objCrypt.GetCurrency: "&objCrypt.GetCurrency()&"<br>")
response.write("objCrypt.GetAmount: "&objCrypt.GetAmount()&"<br>")
response.write("objCrypt.GetShopTransactionID: "&objCrypt.GetShopTransactionID()&"<br>")

call objCrypt.Encrypt()

response.write("objCrypt.GetShopLogin after encrypt: "&objCrypt.GetShopLogin()&"<br>")
response.write("objCrypt.GetEncryptedString after encrypt: "&objCrypt.GetEncryptedString()&"<br>")

response.write("objCrypt.GetErrorCode: "&objCrypt.GetErrorCode()&"<br>")
response.write("objCrypt.GetErrorDescription: "&objCrypt.GetErrorDescription()&"<br>")

if objCrypt.GetErrorCode = 0 then
	b = objCrypt.GetEncryptedString
	a = objCrypt.GetShopLogin
end if

'FINE SCRIPT PER CRITTOGRAFIA.

'SE TUTTO OK SI HANNO 2 VARIABILI A E B DA UTILIZZARE PER IL 'PASSAGGIO DEI PARAMETRI A BANCA SELLA

'ESEMPIO CON FORM HTML
'Per codici test impostare il dominio https://testecomm.sella.it
%>

<form action="https://testecomm.sella.it/gestpay/pagam.asp">
<input name="a" type="hidden" value="<%=a%>">
<input name="b" type="hidden" value="<%=b%>">
<input type="submit" value=" OK " name="Input">
</form>

</body>
</html>