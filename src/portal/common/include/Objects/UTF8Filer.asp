<%
' UTF8Filer 1.4
' Written by Hunter Beanland
' http://www.geocities.com/hbeanland/
' hbeanland@yahoo.com.au

' UTF8Filer is a class which can read and write UTF-8 (Mixed Byte) files being raw text, Delimited (CSV, Tab etc) or Fixed Width data files. 
' It includes functions to convert to and from Unicode as well as some useful mixed byte text handling functions.
' ASP (pre .NET) does not support mixed byte files or text streams - only pure Single Byte and Unicode, so we need use this class to help us out.
' Full Microsoft Excel delimited rules are also adhered to:
' Text including commas & tabs encapsulated in double quotes are treated as 1 field ("text,text" = text,text)
' Double quotes used in encapsulation are stripped ("Text" = Text)
' 2 sets of double quotes are treated as 1 set of double quotes ("" = escaped ")
' Text in quotes can run over several lines
' Fields are outputed in a variable length single dimension array & inputed via the same.

Class UTF8Filer
	Public ErrorText		'Empty if no error, otherwise error text
	Public VirtualFileName	'Virtual Filename
	Public AbsoluteFileName 'Absolute / Physical Filename
	Public LineDelimiter	'vbCRLF = carriage return & line feed, vbLF = Line feed, etc
	Public UnicodeCharset	'Windows-1252, X-ANSI, big5, gb2312, shift_jis, EUC-KR, UTF-8, UTF-7, ASCII, etc
	Public LineNumber		'Current Read/Write line number counter
	Public CharNumber		'Currect Read char number counter
	Public TextBuffer		'Text Buffer for most operations
	Public TextBufferType	'1 = Single Byte, 2 = Unicode, 3 Mixed Byte
	Public Fields			'Data fields array
	Public Big5Space		'Big5 (T.Chinese) space (pseudo-constant) can be used in setting FieldPadding
	Private ColWidths		'Array of fixed width column widths
	Private ColPadding()	'Array of fixed width column padding info
	Private FieldDelimiter	'Delimiter
	Private FileMode		'D = Delimited, F = Wixed Width, T = Text (non-structured)
	Private ReadLineBuffer	'Text Buffer of unread data (forward of CharNumber)

	Public Property Let Delimiter(ByVal strDelimiter)
	'Delimiter to use: , = Comma*, vbTab = Tab, etc (sinlge char only) Not used for fixed width files
		FieldDelimiter = left(strDelimiter,1)
		FileMode = "D"
	End Property

	Public Property Get Delimiter
	'Return the delimiter
		Delimiter = FieldDelimiter
	End Property

	Public Property Let FieldWidths(ByVal strColWidths)
	'Char widths with comma seperator. ie "10,15,4,2,100". Fixed width files only
	Dim ArrayIndex
		ColWidths = split(strColWidths,",")
		for ArrayIndex = 0 to ubound(ColWidths)
			if not isnumeric(ColWidths(ArrayIndex)) then ColWidths(ArrayIndex) = 0
			ColWidths(ArrayIndex) = cint(ColWidths(ArrayIndex))
		next
		if UBound(ColPadding,1) < ArrayIndex then
			redim preserve ColPadding(1,ArrayIndex)
		end if
		FileMode = "F"
	End Property

	Public Property Get FieldWidths
	'Return array of widths
		FieldWidths = ColWidths
	End Property
	
	Public Property Let FieldPadding(ByVal strColPadding)
	'Left/Right + single byte/unicode (not UTF8) char(s) padding with comma separator. ie "R ,,L0,R-,R" & chrw(&H3000). 
	'Default "R " or right space. Writing Fixed width files only
	Dim ArrayIndex, ColPads
		ColPads = split(strColPadding,",")
		redim ColPadding(1,ubound(ColPads))
		for ArrayIndex = 0 to ubound(ColPads)
			if left(ColPads(ArrayIndex),1) <> "L" then 
				ColPadding(0,ArrayIndex) = "R"
			else
				ColPadding(0,ArrayIndex) = "L"
			end if
			if len(ColPads(ArrayIndex)) > 1 then
				ColPadding(1,ArrayIndex) = right(ColPads(ArrayIndex),len(ColPads(ArrayIndex))-1)
			else
				ColPadding(1,ArrayIndex) = " "
			end if	
		next
	End Property

	Public Property Get FieldPadding
	'Return array of padding data
		FieldPadding = ColPadding
	End Property
		
	Private Sub Class_Initialize
		ErrorText = ""
		LineDelimiter = vbCRLF
		UnicodeCharset = "big5"
    	TextBufferType = 2
    	CharNumber = 1
    	LineNumber = 0
		Delimiter = ","
		redim ColPadding(1,1)
		FieldWidths = "1"
		FileMode = "T"
		Big5Space = chrw(&H3000)
	End Sub

	Public Function LoadFile(ByVal FileName)
		'Check file name and open the file
		Dim Stream
		if instr(trim(FileName),":") = 2 then 
			AbsoluteFileName = trim(FileName)
		elseif instr(trim(FileName),".") = 1 then
			ErrorText = "Path must be full virtual or absolute (. and .. not allowed)"
			LoadFile = False
			exit Function
		else
			VirtualFileName = trim(FileName)
			AbsoluteFileName = Server.MapPath(trim(FileName))
		end if
		LineNumber = 0
		'Read text stream in as text (as it is mixed single and double byte) then treat it as binary data (X-ANSI).
		'The stream will need to be converted to the original charset before it can be read.
		Set Stream = Server.CreateObject("ADODB.Stream")
		With Stream
			.Charset = "Windows-1252"
			.Type = 2	'adTypeText
			.Open
			.LoadFromFile AbsoluteFileName
			.Position = 0
			.Charset = "X-ANSI"
			.Type = 1	'adTypeBinary
			'Future version: Need to load smaller sections at a time
			TextBuffer = .Read			
			.Close
		End With
   		TextBufferType = 3
		Set Stream = nothing
		LoadFile = True
	End Function

	Public Function SaveFile(ByVal FileName)
		'Check file name, TextBufferType and save to the file
		Dim Stream
		if instr(trim(FileName),":") = 2 then 
			AbsoluteFileName = trim(FileName)
		elseif instr(trim(FileName),".") = 1 then
			ErrorText = "Path must be full virtual or absolute (. and .. not allowed)"
			SaveFile = False
			exit Function
		else
			VirtualFileName = trim(FileName)
			AbsoluteFileName = Server.MapPath(trim(FileName))
		end if
		Set Stream = Server.CreateObject("ADODB.Stream")
		'Convert and Save in one step. Just taking a UTF8 stream in a VB variant variable gives Stream errors for some reason
		If TextBufferType = 3 then cTextBuffer2Unicode
    	With Stream
    		.Charset = UnicodeCharset
    		.Type = 2	'adTypeText
    		.Open
    		.WriteText TextBuffer
    		.Position = 0
    		.Charset = "X-ANSI"
    		.Type = 1	'adTypeBinary
      		.SaveToFile AbsoluteFileName, 2
     		.Close
    	End With
		Set Stream = nothing
		SaveFile = True
	End Function
		
	Public Sub cTextBuffer2UTF8
    	'Converts TextBuffer from Unicode to Binary (which is UTF-8)
    	'Based on techiniques by Lewis Moten
		Dim Stream
    	if isnull(TextBuffer) or TextBuffer = "" or TextBufferType = 3 Then Exit Sub
		Set Stream = Server.CreateObject("ADODB.Stream")
    	With Stream
    		.Charset = UnicodeCharset
    		.Type = 2	'adTypeText
    		.Open
    		.WriteText TextBuffer
    		.Position = 0
    		.Charset = "X-ANSI"
    		.Type = 1	'adTypeBinary
    		.Position = 0
    		TextBuffer = MidB(.Read, 1)
      		.Close
    	End With
	   	TextBufferType = 3
		Set Stream = Nothing
	End Sub
			
	Public Function cUnicode2UTF8(UnicodeText)
    	'Converts from Unicode to Binary (which is UTF-8)
    	'Based on techiniques by Lewis Moten
		Dim Stream
    	if isnull(UnicodeText) or UnicodeText = "" Then Exit Function
		Set Stream = Server.CreateObject("ADODB.Stream")
    	With Stream
    		.Charset = UnicodeCharset
    		.Type = 2	'adTypeText
    		.Open
    		.WriteText UnicodeText
    		.Position = 0
    		.Charset = "X-ANSI"
    		.Type = 1	'adTypeBinary
    		.Position = 0
    		cUnicode2UTF8 = MidB(.Read, 1)
      		.Close
    	End With
		Set Stream = Nothing
	End Function

	Public Sub cTextBuffer2Unicode
    	'Converts TextBuffer from binary (which is the UTF-8 data from we read from the file) to a Unicode string
    	'Based on techiniques by Lewis Moten & Cakkie
		Dim Stream, Length, Buffer, Rs
    	if isnull(TextBuffer) or len(TextBuffer) = 0 or TextBufferType = 2 Then Exit Sub
		Set Stream = Server.CreateObject("ADODB.Stream")
    	TextBuffer = MidB(TextBuffer, 1)
    	Length = LenB(TextBuffer)
    	Set Rs = Server.CreateObject("ADODB.Recordset")
    	Call Rs.Fields.Append("BinaryData", 205, Length)	'205 = adLongVarBinary
    	Rs.Open
    	Rs.AddNew
    	Rs.Fields("BinaryData").AppendChunk(TextBuffer & ChrB(0))
    	Rs.Update
    	Buffer = Rs.Fields("BinaryData").GetChunk(Length)
    	Rs.Close
    	Set Rs = Nothing
    	Stream.Charset = "X-ANSI"
    	Stream.Type = 1	'adTypeBinary
    	Stream.Open
    	Call Stream.Write(Buffer)
    	Stream.Position = 0
    	Stream.Type = 2	'adTypeText
    	Stream.Charset = UnicodeCharset
    	TextBuffer = Stream.ReadText(-1)
    	TextBufferType = 2
		Stream.Close
		Set Stream = Nothing
	End Sub

	Public Function cUTF82Unicode(UTF8Text)
    	'Converts from binary (which is the UTF-8 data from we read from the file) to a Unicode string
    	'Based on techiniques by Lewis Moten & Cakkie
		Dim Stream, Length, Buffer, Rs
    	if isnull(UTF8Text) or len(UTF8Text) = 0 Then Exit Function
		Set Stream = Server.CreateObject("ADODB.Stream")
    	UTF8Text = MidB(UTF8Text, 1)
    	Length = LenB(UTF8Text)
    	Set Rs = Server.CreateObject("ADODB.Recordset")
    	Call Rs.Fields.Append("BinaryData", 205, Length)	'205 = adLongVarBinary
    	Rs.Open
    	Rs.AddNew
    	Rs.Fields("BinaryData").AppendChunk(UTF8Text & ChrB(0))
    	Rs.Update
    	Buffer = Rs.Fields("BinaryData").GetChunk(Length)
    	Rs.Close
    	Set Rs = Nothing
    	Stream.Charset = "X-ANSI"
    	Stream.Type = 1	'adTypeBinary
    	Stream.Open
    	Call Stream.Write(Buffer)
    	Stream.Position = 0
    	Stream.Type = 2	'adTypeText
    	Stream.Charset = UnicodeCharset
    	cUTF82Unicode = Stream.ReadText(-1)
		Stream.Close
		Set Stream = Nothing
	End Function

	Public Function EOF
		'Check if ReadLine has read up to the End of File (TextBuffer)
		ErrorText = ""
		EOF = False
		if TextBufferType = 3  then		
			if CharNumber >= lenB(TextBuffer) then
				ErrorText = "End of File reached"
				EOF = True
			end if
		else
			if CharNumber >= len(TextBuffer) then
				ErrorText = "End of File reached"
				EOF = True
			end if
		end if
	End Function

	Public Function ReadLine
		'Find the next line of data from the buffer and return it
		Dim LineBuffer, NextEOL
		if TextBufferType = 3  then
			'(Mixed) Byte level
			'NextEOL = InstrMBBuffer(CharNumber,vbLF)
			NextEOL = InstrMBBuffer(CharNumber,chr(&H0A) & chr(&H00))
			'NextEOL = Instr(CharNumber,TextBuffer,chr(&H0A) & chr(&H00),1)
			if NextEOL < 1 then NextEOL = lenB(TextBuffer)
			LineBuffer=midB(TextBuffer,CharNumber,NextEOL-CharNumber)
			if leftB(LineBuffer,1) = vbCR then LineBuffer = rightB(LineBuffer,lenB(LineBuffer)-1)
			if rightB(LineBuffer,1) = vbCR then LineBuffer = leftB(LineBuffer,lenB(LineBuffer)-1)
		else
			'Unicode Char level
			NextEOL = instr(CharNumber,TextBuffer,vbLF)
			if NextEOL < 1 then NextEOL = len(TextBuffer)
			LineBuffer=mid(TextBuffer,CharNumber,NextEOL-CharNumber)
			if left(LineBuffer,1) = vbCR then LineBuffer = right(LineBuffer,len(LineBuffer)-1)
			if right(LineBuffer,1) = vbCR then LineBuffer = left(LineBuffer,len(LineBuffer)-1)
		end if
		CharNumber = NextEOL + 1
		LineNumber = LineNumber + 1
		If FileMode = "D" then
			SplitDelimiter(LineBuffer)
			ReadLine = Fields
		Elseif FileMode = "F" then
			call SplitFixed(LineBuffer)
			ReadLine = Fields
		Else
			ReadLine = LineBuffer
		End if
	End Function

	Public Function WriteLine(ByVal FieldArray)
		'Write a line on the end of the buffer (Usually start with a clear buffer)
		Dim LineSource, intLength, ArrayIndex
		LineSource = ""
		If ((FileMode = "D" or FileMode = "F") and not isarray(FieldArray)) or (FileMode = "T" and isarray(FieldArray)) then
				ErrorText = "Supplied Data is not the correct type for this file"
				WriteLine = False
				Exit Function		
		end if
		If FileMode = "D" then
			'Delimited
			for i = 0 to ubound(FieldArray,1)
				if isnull(FieldArray(i)) then FieldArray(i) = ""
				if LineSource <> "" then LineSource = LineSource & Delimiter
				if instr(1,FieldArray(i),"""") > 0 or instr(1,FieldArray(i),vbLF) >0 or instr(1,FieldArray(i),Delimiter) >0 then
					LineSource = LineSource & """" & FieldArray(i) & """"
				else
					LineSource = LineSource & FieldArray(i)
				end if	
			Next
			TextBuffer = TextBuffer & LineSource & LineDelimiter
		Elseif FileMode = "F" then
			'Fixed Width
			if ubound(FieldArray,1) <> ubound(ColWidths,1) then
				ErrorText = "Size of array is not the same as the number of column widths supplied"
				WriteLine = False
				Exit Function		
			end if
			for ArrayIndex = 0 to ubound(FieldArray,1)
				if isnull(FieldArray(ArrayIndex)) then FieldArray(ArrayIndex) = ""
				intLength =  LenMB(FieldArray(ArrayIndex))
				if intLength > ColWidths(ArrayIndex) then 
					LineSource = LineSource & leftMB(FieldArray(ArrayIndex),ColWidths(ArrayIndex))
				elseif intLength < ColWidths(ArrayIndex) then
					if ColPadding(0,ArrayIndex) = "L" then
						LineSource = LineSource & StringRepeat(ColWidths(ArrayIndex)-intLength,ColPadding(1,ArrayIndex)) & FieldArray(ArrayIndex)
					else
						LineSource = LineSource & FieldArray(ArrayIndex) & StringRepeat(ColWidths(ArrayIndex)-intLength,ColPadding(1,ArrayIndex))
					end if	
				else
					LineSource = LineSource & FieldArray(ArrayIndex)
				end if	
			Next
			TextBuffer = TextBuffer & LineSource & LineDelimiter
		Else
			'Unstructured Text
			TextBuffer = TextBuffer & FieldArray & LineDelimiter
		End if	
		LineNumber = LineNumber + 1
		WriteLine = True
	End Function

	Public Function SplitDelimiter(SourceLine)
	'Take the line from the source string and split it into an array. 
	Dim SourceIndex, SourceChar, OutputArray(), OutputField, OutputIndex, InQuotes
		InQuotes = false
		OutputIndex = 0
		OutputField = ""
		if TextBufferType = 3  then
			for SourceIndex = 1 to lenB(SourceLine)
				SourceChar = midB(SourceLine, SourceIndex, 1)
				if SourceChar = """" and InQuotes then 
					InQuotes = false
				elseif SourceChar = """" and not InQuotes then 
					InQuotes = true
				end if	
				if (SourceChar = FieldDelimiter and not InQuotes) or SourceIndex = lenB(SourceLine) then 
					if SourceIndex = lenB(SourceLine) and SourceChar <> FieldDelimiter then OutputField = OutputField & SourceChar
					if leftB(OutputField,1) = """" then OutputField = rightB(OutputField,lenB(OutputField)-1)
					if rightB(OutputField,1) = """" then OutputField = leftB(OutputField,lenB(OutputField)-1)
					OutputField = replace(OutputField,"""""","""")
					redim preserve OutputArray(OutputIndex)
					OutputArray(OutputIndex) = OutputField
					OutputField = ""
					OutputIndex = OutputIndex + 1
				else
					OutputField = OutputField & SourceChar
				end if
			next	
		else
			for SourceIndex = 1 to len(SourceLine)
				SourceChar = mid(SourceLine, SourceIndex, 1)
				if SourceChar = """" and InQuotes then 
					InQuotes = false
				elseif SourceChar = """" and not InQuotes then 
					InQuotes = true
				end if	
				if (SourceChar = FieldDelimiter and not InQuotes) or SourceIndex = len(SourceLine) then 
					if SourceIndex = len(SourceLine) and SourceChar <> FieldDelimiter then OutputField = OutputField & SourceChar
					if left(OutputField,1) = """" then OutputField = right(OutputField,len(OutputField)-1)
					if right(OutputField,1) = """" then OutputField = left(OutputField,len(OutputField)-1)
					OutputField = replace(OutputField,"""""","""")
					redim preserve OutputArray(OutputIndex)
					OutputArray(OutputIndex) = OutputField
					OutputField = ""
					OutputIndex = OutputIndex + 1
				else
					OutputField = OutputField & SourceChar
				end if
			next
		end if	
		Fields = OutputArray
	End Function
	
	Public Function SplitFixed(SourceLine)
	'Chop the string into the Fields array with the sizes given in FieldWidths
	Dim SourceIndex, OutputArray(), OutputIndex
		SourceIndex = 1
		for OutputIndex = 0 to ubound(ColWidths,1)
			redim preserve OutputArray(OutputIndex)
			OutputArray(OutputIndex) = midB(SourceLine,SourceIndex,ColWidths(OutputIndex))
			SourceIndex = SourceIndex + ColWidths(OutputIndex)
		next
		Fields = OutputArray
	End Function
	
	Private Function CountChar(SourceLine, SearchChar)
	'Counts the number of times SearchChar occurs in SourceLine. Used by SplitDelimiter
		Dim SourceIndex, HitCount
		HitCount = 0
		for SourceIndex = 1 to lenB(SourceLine)
			if midB(SourceLine,SourceIndex,1) = SearchChar then HitCount = HitCount +1
		next
		CountChar = HitCount
	End Function
	
	Public Function InstrMB(StartChar, UTF8Text, ByVal SearchChar)
	'Search for char in Mixed Byte string. Instr in binary mode doesn't seem to work with a binary string array.
		Dim SourceIndex
		SearchChar = ascb(SearchChar)
		for SourceIndex = StartChar to lenB(UTF8Text)
			if ascb(midB(UTF8Text,SourceIndex,1)) = SearchChar then 
				InstrMB = SourceIndex
				exit Function
			end if
		next
		InstrMB = 0
	End Function	
	
	Public Function InstrMBBuffer(StartChar, ByVal SearchChar)
	'Search for char in Mixed Byte buffer. Instr in binary mode doesn't seem to work with a binary string array.
	'WARNING: This function can be slow for large files (>2MB). Any ideas apart from using ASP.NET ?
		Dim SourceIndex
		'SearchChar = ascb(SearchChar)
		'response.Write("1:" & timer & "<br>")
		for SourceIndex = StartChar to lenB(TextBuffer)
			if midB(TextBuffer,SourceIndex,1) = SearchChar then 
				InstrMBBuffer = SourceIndex
				exit Function
			end if
		next
		'response.Write("2:" & timer & "<br>")
		InstrMBBuffer = 0
	End Function
	
	Public Function InstrMBBuffer2(StartChar, ByVal SearchChar)
		'Search for char in Mixed Byte string. Instr in binary mode doesn't seem to work with a binary string array.
		'Experimental: Does not work.
		Dim regEx, Matches, Match, regExBuffer
		Set regEx = New RegExp
		regEx.Pattern = "\x0A"
		regEx.IgnoreCase = false
		regEx.Global = false
		regExBuffer = rightB(TextBuffer,len(TextBuffer)-CharNumber)
		Set Matches = regEx.Execute(regExBuffer)
		InstrMBBuffer = 0
		response.Write("#@#")
		For Each Match in Matches
		response.Write("#" &  Match.FirstIndex & "#")
			InstrMBBuffer = Match.FirstIndex + CharNumber
		next
	End Function	
	
	Public Function LeftMB(strMixed,intChars)
		'Return left # of UTF-8 (mixed byte) chars in a Unicode stream
		'Double byte char starts with null or high ascii code
		'Single byte char starts with low ascii code (not null) and is proceeded with a null
		Dim i, intCount, IsFirstByte, intAsc, strOutput
		intCount =0
		IsFirstByte = true
		for i = 1 to lenB(strMixed)
			intAsc = ascb(midb(strMixed,i,1))
			if IsFirstByte and i < lenB(strMixed) then
				IsFirstByte = false		
			else
				if intAsc = 0 then 
					intCount = intCount +1
				else
					intCount = intCount +2
				end if
				if intCount <= intChars then strOutput = strOutput & midb(strMixed,i-1,2)
				IsFirstByte = true
			end if	
		next
		LeftMB = strOutput
	End Function

	Public Function LenMB(strMixed)
		'Count UTF-8 (mixed byte) chars in a Unicode stream
		'Double byte char starts with null or high ascii code
		'Single byte char starts with low ascii code (not null) and is proceeded with a null
		Dim i, intCount, IsFirstByte, intAsc
		intCount =0
		IsFirstByte = true
		for i = 1 to lenB(strMixed)
			intAsc = ascb(midb(strMixed,i,1))
			if IsFirstByte and i < lenB(strMixed) then
				IsFirstByte = false		
			else
				if intAsc = 0 then 
					intCount = intCount +1
				else
					intCount = intCount +2
				end if
				IsFirstByte = true
			end if	
		next
		LenMB = intCount
	End Function
	
	Public Function StringRepeat(intLength,strChars)
		'Insert multiple strChars (ie dbl byte chars) to fill intLength bytes
		'VBScript String function can only use single chars
		if LenMB(strChars) = 2 then 
			for i = 1 to intLength step lenMB(strChars)
				StringRepeat = StringRepeat & strChars
			next
		else
			if strChars = "" then strChars = " "
			StringRepeat = string(intLength,strChars)
		end if	
	End Function

	Private Sub Class_Terminate
		'Make sure file is closed (to release file locks and resources)
	End Sub	
End Class

%>
