<%
Class GUIDClass
	Function CreateGUIDTime()
	  Dim tmpTemp
	  tmpTemp = Right(String(4,48) & Year(Now()),4)
	  tmpTemp = tmpTemp & Right(String(4,48) & Month(Now()),2)
	  tmpTemp = tmpTemp & Right(String(4,48) & Day(Now()),2)
	  tmpTemp = tmpTemp & Right(String(4,48) & Hour(Now()),2)
	  tmpTemp = tmpTemp & Right(String(4,48) & Minute(Now()),2)
	  tmpTemp = tmpTemp & Right(String(4,48) & Second(Now()),2)
	  CreateGUIDTime = tmpTemp
	End Function
	
	Function CreateGUIDTime2()
	  Dim tmpTemp1,tmpTemp2
	  tmpTemp1 = Right(String(15,48) & CStr(CLng(DateDiff("s","1/1/2000",Date()))), 15)
	  tmpTemp2 = Right(String(5,48) & CStr(CLng(DateDiff("s","12:00:00 AM",Time()))), 5)
	  CreateGUIDTime2 = tmpTemp1 & tmpTemp2
	End Function
	
	Function CreateGUIDTime3()
	  Randomize Timer
	  Dim tmpTemp1,tmpTemp2,tmpTemp3
	  tmpTemp1 = Right(String(10,48) & CStr(CLng(DateDiff("s","1/1/2000",Date()))), 10)
	  tmpTemp2 = Right(String(5,48) & CStr(CLng(DateDiff("s","12:00:00 AM",Time()))), 5)
	  tmpTemp3 = Right(String(5,48) & CStr(Int(Rnd(1) * 100000)),5)
	  CreateGUIDTime3 = tmpTemp1 & tmpTemp2 & tmpTemp3
	End Function
	
	Function CreateGUIDTime4()
	  Randomize Timer
	  Dim tmpTemp1,tmpTemp2,tmpTemp3
	  tmpTemp1 = Right(String(10,48) & CStr(CLng(DateDiff("s","1/1/2000",Date()))), 7)
	  tmpTemp2 = Right(String(5,48) & CStr(CLng(DateDiff("s","12:00:00 AM",Time()))), 3)
	  CreateGUIDTime4 = tmpTemp1 & tmpTemp2
	End Function
	
	Function CreateGUIDRandom()
	  Randomize Timer
	  Dim tmpCounter,tmpGUID
	  Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	  For tmpCounter = 1 To 20
		tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
	  Next
	  CreateGUIDRandom = tmpGUID
	End Function
	
	Function CreateGUIDRandomVarLenght(tmpLength)
	  Randomize Timer
	  Dim tmpCounter,tmpGUID
	  Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	  For tmpCounter = 1 To tmpLength
		tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
	  Next
	  CreateGUIDRandomVarLenght = tmpGUID
	End Function
	
	Function CreateNumberGUIDRandomVarLenght(tmpLength)
	  Randomize Timer
	  Dim tmpCounter,tmpGUID
	  Const strValid = "0123456789"
	  For tmpCounter = 1 To tmpLength
		tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
	  Next
	  CreateNumberGUIDRandomVarLenght = tmpGUID
	End Function
	
	Function CreateWindowsGUID()
	  CreateWindowsGUID = CreateGUIDRandomVarLenght(8) & "-" & _
		CreateGUIDRandomVarLenght(4) & "-" & _
		CreateGUIDRandomVarLenght(4) & "-" & _
		CreateGUIDRandomVarLenght(4) & "-" & _
		CreateGUIDRandomVarLenght(12)
	End Function
	
	Function CreateOrderGUIDLong()
	  	CreateOrderGUIDLong = CreateGUIDTime3() & _
	    CreateGUIDRandomVarLenght(5) & _
		CreateGUIDRandomVarLenght(10) & _
		CreateGUIDRandomVarLenght(15) & _
		CreateGUIDRandomVarLenght(20) & _
		CreateGUIDRandomVarLenght(25) & _
		CreateGUIDRandomVarLenght(50)
	End Function
	
	Function CreateOrderGUID()
	  CreateOrderGUID = CreateGUIDTime3()' & _
		'CreateGUIDRandomVarLenght(10) & _
		'CreateGUIDRandomVarLenght(15) & _
		'CreateGUIDRandomVarLenght(25) & _
		'CreateGUIDRandomVarLenght(25)
	End Function
	
	Function CreateUserGUID()
	  CreateUserGUID = CreateGUIDTime3() & _
		CreateGUIDRandomVarLenght(10) & _
		CreateGUIDRandomVarLenght(15) & _
		CreateGUIDRandomVarLenght(25)
	End Function
	
	Function CreatePasswordGUID()
	  CreatePasswordGUID = CreateGUIDRandomVarLenght(12)
	End Function
	
	Function CreateVoucherCodeGUID()
	  CreateVoucherCodeGUID = CreateGUIDTime3()' & _
		CreateGUIDRandomVarLenght(10)
	End Function
	
	Function orderGUIDLength
		 orderGUIDLength = CInt(100)
	End Function
End Class
%>