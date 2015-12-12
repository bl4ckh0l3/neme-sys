<%
Class UtilClass	
	Public Function getUniqueKeyOrderIdPayment()
		getUniqueKeyOrderIdPayment = "id_order_ack"
	End Function
	
	Public Function getUniqueKeyOrderAmountPayment()
		getUniqueKeyOrderAmountPayment = "amount_order_ack"
	End Function
	
	Public Function getUniqueKeyOrderGUIDPayment()
		getUniqueKeyOrderGUIDPayment = "order_guid_ack"
	End Function
	
	Public Function getUniqueKeyOrderTypePayment()
		getUniqueKeyOrderTypePayment = "payment_type_ack"
	End Function
	
	Public Function getUniqueKeyExtURLPayment()
		getUniqueKeyExtURLPayment = "external_url"
	End Function	
	
	Public Function getUniqueKeySuccessPaymentTransaction()
		getUniqueKeySuccessPaymentTransaction = "SUCCESS"
	End Function	
	
	Public Function getUniqueKeyPendingPaymentTransaction()
		getUniqueKeyPendingPaymentTransaction = "PENDING"
	End Function
	
	Public Function getUniqueKeyFailedPaymentTransaction()
		getUniqueKeyFailedPaymentTransaction = "FAILED"
	End Function
	
	Public Function getUniqueKeyEncryptDecrypt()
		getUniqueKeyEncryptDecrypt = "$bl4ckh0l3$"
	End Function
	
	'Function URLDecode(str) 
	'	str = Replace(str, "+", " ") 
	'	For i = 1 To Len(str) 
	'		sT = Mid(str, i, 1) 
	'		If sT = "%" Then 
	'			If i+2 < Len(str) Then 
	'				sR = sR & _ 
	'					Chr(CLng("&H" & Mid(str, i+1, 2))) 
	'				i = i+2 
	'			End If 
	'		Else 
	'			sR = sR & sT 
	'		End If 
	'	Next 
	'	URLDecode = sR 
	'End Function 	
	
	Function URLDecode(sConvert)
		Dim aSplit
		Dim sOutput
		Dim I
		If IsNull(sConvert) Then
		URLDecode = ""
		Exit Function
		End If

		' convert all pluses to spaces
		sOutput = REPLACE(sConvert, "+", " ")

		' next convert %hexdigits to the character
		aSplit = Split(sOutput, "%")

		If IsArray(aSplit) Then
			sOutput = aSplit(0)
			For I = 0 to UBound(aSplit) - 1
				sOutput = sOutput & _
				Chr("&H" & Left(aSplit(i + 1), 2)) &_
				Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
			Next
		End If

		URLDecode = sOutput
	End Function
	
	Function URLEncode(str) 
		URLEncode = Server.URLEncode(str) 
	End Function 

	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = doubleValue
		
		'if (Application("dbType") = 0) then
			convertDoubleDelimiter = Replace(convertDoubleDelimiter, ".",",")
		'else		
			'convertDoubleDelimiter = Replace(convertDoubleDelimiter, ",",".")
		'end if			
	End Function

	Public Function convertDoubleDelimiter4External(doubleValue)
		convertDoubleDelimiter4External = Replace(doubleValue, ".","")	
		convertDoubleDelimiter4External = Replace(doubleValue, ",",".")		
	End Function
End Class
%>