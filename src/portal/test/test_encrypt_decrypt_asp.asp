<%
'******************************
'   KeyGeN.asp
'******************************

Const g_KeyLocation = "F:\Lavoro\Portal Project\demo_portal\portal\test\test_key.txt"
Const g_KeyLen = 256

On Error Resume Next

Call WriteKeyToFile(KeyGeN(g_KeyLen),g_KeyLocation)

if Err <> 0 Then
   Response.Write "ERROR GENERATING KEY" & "<P>"
   Response.Write Err.Number & "<BR>"
   Response.Write Err.Description & "<BR>"  
Else
   Response.Write "KEY SUCCESSFULLY GENERATED"
End If


Sub WriteKeyToFile(MyKeyString,strFileName)
   Dim keyFile, fso
   set fso = Server.CreateObject("scripting.FileSystemObject") 
   set keyFile = fso.CreateTextFile(strFileName, true) 
   keyFile.WriteLine(MyKeyString)
   keyFile.Close
End Sub

Function KeyGeN(iKeyLength)
Dim k, iCount, strMyKey
   lowerbound = 35 ' 35
   upperbound = 96 ' 96
   Randomize      ' Initialize random-number generator.
   for i = 1 to iKeyLength
      s = 255
      k = Int(((upperbound - lowerbound) + 1) * Rnd + lowerbound)
      strMyKey =  strMyKey & Chr(k) & ""
   next
   KeyGeN = strMyKey
End Function
%>


<%
'******************************
'   Crypt.asp
'******************************

Dim g_Key

Const g_CryptThis = "0000002911680000562743178CGN5EFOLRO5NB3ZW026MOHPYCXC85W194ZD2JKHIXDAIZHJB2RBPY7ZRYZF93XW6TMHM1E1N0M6MEW8AYN0EM0AT3RJIP6WVKTF4FN4ZDBPO5S42D7V6BUUVPJCSX|85|16.00"

'g_Key = mid(ReadKeyFromFile(g_KeyLocation),1,Len(g_CryptThis))

g_Key = KeyGeN(g_KeyLen)

Response.Write "<p>ORIGINAL STRING: " & g_CryptThis & "<p>"
Response.Write "<p>KEY VALUE: " & g_Key  & "<p>"
Response.Write "<p>ENCRYPTED CYPHERTEXT: " & EnCrypt(g_CryptThis) & "<p>"
Response.Write "<p>DECRYPTED CYPHERTEXT: " & DeCrypt(EnCrypt(g_CryptThis)) & "<p>"

Function EnCrypt(strCryptThis)
   Dim strChar, iKeyChar, iStringChar, i
   for i = 1 to Len(strCryptThis)
      iKeyChar = Asc(mid(g_Key,i,1))
      iStringChar = Asc(mid(strCryptThis,i,1))
      ' *** uncomment below to encrypt with addition,
      ' iCryptChar = iStringChar + iKeyChar
      iCryptChar = iKeyChar Xor iStringChar
      strEncrypted =  strEncrypted & Chr(iCryptChar)
   next
   EnCrypt = strEncrypted
End Function

Function DeCrypt(strEncrypted)
Dim strChar, iKeyChar, iStringChar, i
   for i = 1 to Len(strEncrypted)
      iKeyChar = (Asc(mid(g_Key,i,1)))
      iStringChar = Asc(mid(strEncrypted,i,1))
      ' *** uncomment below to decrypt with subtraction	
      ' iDeCryptChar = iStringChar - iKeyChar 
      iDeCryptChar = iKeyChar Xor iStringChar
      strDecrypted =  strDecrypted & Chr(iDeCryptChar)
   next
   DeCrypt = strDecrypted
End Function

Function ReadKeyFromFile(strFileName)
   Dim keyFile, fso, f
   set fso = Server.CreateObject("Scripting.FileSystemObject") 
   set f = fso.GetFile(strFileName) 
   set ts = f.OpenAsTextStream(1, -2)

   Do While not ts.AtEndOfStream
     keyFile = keyFile & ts.ReadLine
   Loop 

   ReadKeyFromFile =  keyFile
End Function
%>


