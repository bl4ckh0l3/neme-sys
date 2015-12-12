<%
Class LocalizationClass
	Private id
	Private id_element
	Private type_elem	
	Private latitude
	Private longitude
	Private txtinfo
	
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub		
	
	Public Function getElemID()
		getElemID = id_element
	End Function
	
	Public Sub setElemID(strElID)
		id_element = strElID
	End Sub	
	
	Public Function getType()
		getType = type_elem
	End Function
	
	Public Sub setType(strType)
		type_elem = strType
	End Sub
	
	Public Function getLatitude()
		getLatitude = latitude
	End Function
	
	Public Sub setLatitude(strLatitude)
		latitude = CDbl(convertDoubleDelimiter(strLatitude))
	End Sub
	
	Public Function getLongitude()
		getLongitude = longitude
	End Function
	
	Public Sub setLongitude(strLongitude)
		longitude = CDbl(convertDoubleDelimiter(strLongitude))
	End Sub
	
	Public Function getInfo()
		getInfo = txtinfo
	End Function
	
	Public Sub setInfo(strInfo)
		txtinfo = strInfo
	End Sub
			
	Public Function insertPoint(strID, strType, strLatitude, strLongitude, strInfo, objConn)
		on error resume next		
		Dim strSQL
		insertPoint=-1		
		
		strSQL = "INSERT INTO googlemap_localization(id_element, `type`, latitude, longitude, txtinfo) VALUES("
		strSQL = strSQL & "?,?,"
		if (isNull(Trim(strLatitude)) OR Trim(strLatitude) = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		if (isNull(Trim(strLongitude)) OR Trim(strLongitude) = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		if (isNull(Trim(strInfo)) OR Trim(strInfo) = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ");"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strID)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,strType)
		if not isNull(Trim(strLatitude)) AND not(Trim(strLatitude) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLatitude)))
		end if
		if not isNull(Trim(strLongitude)) AND not(Trim(strLongitude) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLongitude)))
		end if
		if not isNull(Trim(strInfo)) AND not(Trim(strInfo) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,Trim(strInfo))
		end if
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(googlemap_localization.id) as id FROM googlemap_localization")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertPoint = objRS("id")	
		end if		
		Set objRS = Nothing

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
			
	Public Function insertPointNoTransaction(strID, strType, strLatitude, strLongitude, strInfo)
		on error resume next		
		Dim strSQL, objConn, objDB
		insertPointNoTransaction=-1	
		
		strSQL = "INSERT INTO googlemap_localization(id_element, `type`, latitude, longitude, txtinfo) VALUES("
		strSQL = strSQL & "?,?,"
		if (isNull(Trim(strLatitude)) OR Trim(strLatitude) = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		if (isNull(Trim(strLongitude)) OR Trim(strLongitude) = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		if (isNull(Trim(strInfo)) OR Trim(strInfo) = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ");"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strID)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,strType)
		if not isNull(Trim(strLatitude)) AND not(Trim(strLatitude) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLatitude)))
		end if
		if not isNull(Trim(strLongitude)) AND not(Trim(strLongitude) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLongitude)))
		end if
		if not isNull(Trim(strInfo)) AND not(Trim(strInfo) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,Trim(strInfo))
		end if
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(googlemap_localization.id) as id FROM googlemap_localization")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertPointNoTransaction = objRS("id")	
		end if		
		Set objRS = Nothing	
		Set objDB = Nothing	
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyPoint(strID, strIdElem, strLatitude, strLongitude, strInfo, objConn)
		on error resume next
		Dim strSQL

		strSQL = "UPDATE googlemap_localization SET id_element=?,"
		if (isNull(Trim(strLatitude)) OR Trim(strLatitude) = "") then
			strSQL = strSQL & "latitude=NULL,"
		else
			strSQL = strSQL & "latitude=?,"
		end if
		if (isNull(Trim(strLongitude)) OR Trim(strLongitude) = "") then
			strSQL = strSQL & "longitude=NULL,"
		else
			strSQL = strSQL & "longitude=?,"
		end if
		if (isNull(Trim(strInfo)) OR Trim(strInfo) = "") then
			strSQL = strSQL & "txtinfo=NULL"
		else
			strSQL = strSQL & "txtinfo=?"
		end if
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strIdElem)
		if not isNull(Trim(strLatitude)) AND not(Trim(strLatitude) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLatitude)))
		end if
		if not isNull(Trim(strLongitude)) AND not(Trim(strLongitude) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLongitude)))
		end if
		if not isNull(Trim(strInfo)) AND not(Trim(strInfo) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,Trim(strInfo))
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strID)
		objCommand.Execute()
		Set objCommand = Nothing

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyPointNoTransaction(strID, strIdElem, strLatitude, strLongitude, strInfo)
		on error resume next
		Dim strSQL, objConn, objDB

		strSQL = "UPDATE googlemap_localization SET id_element=?,"
		if (isNull(Trim(strLatitude)) OR Trim(strLatitude) = "") then
			strSQL = strSQL & "latitude=NULL,"
		else
			strSQL = strSQL & "latitude=?,"
		end if
		if (isNull(Trim(strLongitude)) OR Trim(strLongitude) = "") then
			strSQL = strSQL & "longitude=NULL,"
		else
			strSQL = strSQL & "longitude=?,"
		end if
		if (isNull(Trim(strInfo)) OR Trim(strInfo) = "") then
			strSQL = strSQL & "txtinfo=NULL"
		else
			strSQL = strSQL & "txtinfo=?"
		end if
		strSQL = strSQL & " WHERE id=?;"
	
'response.Write(strSQL)
	
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strIdElem)
		if not isNull(Trim(strLatitude)) AND not(Trim(strLatitude) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLatitude)))
		end if
		if not isNull(Trim(strLongitude)) AND not(Trim(strLongitude) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLongitude)))
		end if
		if not isNull(Trim(strInfo)) AND not(Trim(strInfo) = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,Trim(strInfo))
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strID)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deletePoint(strID)
		on error resume next
		Dim objDB, strSQLDel, objConn
		
		strSQLDel = "DELETE FROM googlemap_localization WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQLDel
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,strID)
		objCommand.Execute()
		Set objCommand = Nothing				
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub	

	Public Function findPointByElement(strID, strType)
		on error resume next		
				
		findPointByElement = null				
		Dim objDB, strSQL, objRS, objConn
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		strSQL = "SELECT * FROM googlemap_localization WHERE id_element=? AND `type`=? ORDER BY id;"	

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strID)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,strType)
		Set objRS = objCommand.Execute()

		if not(objRS.EOF) then
			Dim objLoc, objListaPoint			
			Set objListaPoint = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF			
				Set objLoc = new LocalizationClass	
				id = objRS("id")
				objLoc.setID(id)
				objLoc.setElemID(objRS("id_element"))
				objLoc.setType(objRS("type"))
				objLoc.setLatitude(objRS("latitude"))
				objLoc.setLongitude(objRS("longitude"))
				objLoc.setInfo(objRS("txtinfo"))				
				objListaPoint.add id, objLoc
				Set objLoc = Nothing
				objRS.moveNext()				
			loop		
						
			Set findPointByElement = objListaPoint			
			Set objListaPoint = nothing
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	

	Public Function findPointByElements(listID, strType)
		on error resume next		
				
		findPointByElements = null				
		Dim objDB, strSQL, objRS, objConn
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
		
		strSQL = "SELECT * FROM googlemap_localization WHERE `type`=?"			
		if not(isNull(listID)) then
			strSQL = strSQL & " AND id_element IN("&listID&")"
		end if
		strSQL = Trim(strSQL)
		strSQL = strSQL & ";"		
		'response.write(strSQL&"<br>")

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,strType)
		Set objRS = objCommand.Execute()

		if not(objRS.EOF) then
			Dim objLoc, objListaPoint, tmpcounter			
			Set objListaPoint = Server.CreateObject("Scripting.Dictionary")
			
			do while not objRS.EOF			
				Set objLoc = new LocalizationClass	
				id = objRS("id")
				objLoc.setID(id)
				objLoc.setElemID(objRS("id_element"))
				objLoc.setType(objRS("type"))
				objLoc.setLatitude(objRS("latitude"))
				objLoc.setLongitude(objRS("longitude"))
				objLoc.setInfo(objRS("txtinfo"))				
				objListaPoint.add id, objLoc
				Set objLoc = Nothing
				objRS.moveNext()				
			loop		
						
			Set findPointByElements = objListaPoint			
			Set objListaPoint = nothing
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function									
		
			
	Public Function findPointByID(strID)
		on error resume next
		
		Set findPointByID = null				
		Dim objDB, strSQL, objRS, objConn, objLoc

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
					
		strSQL = "SELECT * FROM googlemap_localization WHERE id=?;"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strID)
		Set objRS = objCommand.Execute()

		if not(objRS.EOF) then					
			Set objLoc = new LocalizationClass					
			objLoc.setID(objRS("id"))				
			objLoc.setElemID(objRS("id_element"))
			objLoc.setType(objRS("type"))
			objLoc.setLatitude(objRS("latitude"))
			objLoc.setLongitude(objRS("longitude"))
			objLoc.setInfo(objRS("txtinfo"))
			Set findPointByID = objLoc				
			Set objLoc = nothing
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function									
		
			
	Public Function findPointByPosition(strID, strType, strLatitude, strLongitude)
		on error resume next
		
		Set findPointByPosition = null				
		Dim objDB, strSQL, objRS, objConn, objLoc

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()
					
		strSQL = "SELECT * FROM googlemap_localization WHERE id_element=? AND `type`=? AND latitude=? AND longitude=?;"
		strSQL = Trim(strSQL)
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,strID)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,strType)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLatitude)))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(Trim(strLongitude)))
		Set objRS = objCommand.Execute()

		if not(objRS.EOF) then					
			Set objLoc = new LocalizationClass					
			objLoc.setID(objRS("id"))				
			objLoc.setElemID(objRS("id_element"))
			objLoc.setType(objRS("type"))
			objLoc.setLatitude(objRS("latitude"))
			objLoc.setLongitude(objRS("longitude"))
			objLoc.setInfo(objRS("txtinfo"))
			Set findPointByPosition = objLoc				
			Set objLoc = nothing
		end if
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function


	'****************************** START: COPPIA DI FUNZIONI STATICHE PER DETERMINARE SE UN PUNTO APPARTIENE AD UN POLIGONO DATO IL PUNTO E I VERTICI DEL POLIGONO
	Private function isLeft(P0, P1, P2 )
		isLeft =  ( (P1.getLatitude() - P0.getLatitude()) * (P2.getLongitude() - P0.getLongitude()) - (P2.getLatitude() -  P0.getLatitude()) * (P1.getLongitude() - P0.getLongitude()) )
	End function

	Public function isPointInPolygon( point, Vertices)
	
		'for each j in Vertices
		'	response.write("lat:"&j.getLatitude()&" - lng:"&j.getLongitude()&"<br>")
		'next
	
		wn = 0    ' the  winding number counter
		keys=Vertices.Keys
		'loop through all edges of the polygon
		for i=0 to Vertices.Count-2    ' edge from V[i] to  V[i+1]
			if (keys(i).getLongitude() <= point.getLongitude()) then          ' start y <= P.y
				if (keys(i+1).getLongitude()  > point.getLongitude()) then      ' an upward crossing
					if (isLeft( keys(i), keys(i+1), point) > 0) then ' P left of  edge
						wn=wn+1            ' have  a valid up intersect
					end if
				end if
			else                         ' start y > P.y (no test needed)
				if (keys(i+1).getLongitude()  <= point.getLongitude()) then     ' a downward crossing
					if (isLeft( keys(i), keys(i+1), point) < 0) then  ' P right of  edge
						wn = wn-1            ' have  a valid down intersect
					end if
				end if
			end if
		Next

		if(wn=0)then
			isPointInPolygon = false
		else
			isPointInPolygon = true
		end if
	End function
	'****************************** END: COPPIA DI FUNZIONI STATICHE PER DETERMINARE SE UN PUNTO APPARTIENE AD UN POLIGONO DATO IL PUNTO E I VERTICI DEL POLIGONO
	

	'****************************** START: FUNZIONE STATICA PER DETERMINARE SE UN PUNTO APPARTIENE AD UN CERCHIO, DATO IL PUNTO, IL CENTRO DEL CERCHIO E IL RAGGIO (il raggio va passato in metri)
	Public function IsPointInCircleOnEarthSurface(punto, center, radiusParam)
		distanceInMeters = greatCircleDistanceInMeters(punto.getLongitude(), punto.getLatitude(), center.getLongitude(), center.getLatitude())		
		bolDiff = (distanceInMeters < CDbl(radiusParam))
		'response.write("distanceInMeters:"&distanceInMeters&" -typename:"&typename(distanceInMeters)&" -radiusParam:"&radiusParam&" -typename:"&typename(radiusParam)&" -bolDiff:"& bolDiff &"<br>")

		if (bolDiff) then
			IsPointInCircleOnEarthSurface=true
		else
			IsPointInCircleOnEarthSurface=false
		end if
	end Function
	'****************************** END: FUNZIONE STATICA PER DETERMINARE SE UN PUNTO APPARTIENE AD UN CERCHIO, DATO IL PUNTO, IL CENTRO DEL CERCHIO E IL RAGGIO (il raggio va passato in metri)



	'****************************** START: FUNZIONE STATICA PER CONVERTIRE UNA STRINGA DI VERTICI ORDINATI SEPARATI DA | E RAPPRESENTANTI UN POLIGONO CHIUSO(CONVAVO E/O CONVESSO) IN UNA LISTA DI OGGETTI LocalizationClass
	'****************************** PER FUNZIONARE CORRETTAMENTE L ULTIMO VERTICE DELLA LISTA DEVE ESSERE LA RIPETIZIONE DEL PRIMO IN MODO DA DEFINIRE CON SICUREZZA IL POLIGONO CHIUSO
	Public Function convertVertices(vertices)
		'response.write("<br>vertices:"&vertices&"<br>")
		listVertices = Split(vertices, "|", -1, 1)	
		if(isArray(listVertices)) then
			Set objListVertices = Server.CreateObject("Scripting.Dictionary")
			'firstV =""
			For y=LBound(listVertices) to UBound(listVertices)
				'response.write("listVertices(y):"&listVertices(y)&"<br>")
				arrLatLon = Split(listVertices(y), ",", -1, 1)
				'if(y=LBound(listVertices))then
				'	firstV = arrLatLon
				'end if
				if(isArray(arrLatLon)) then
					Set pointV = new LocalizationClass
					pointV.setLatitude(arrLatLon(0))
					pointV.setLongitude(arrLatLon(1))	
					objListVertices.add pointV, ""		
					Set pointV = nothing
				end if
			next
			'if(isArray(firstV)) then
			'	Set pointV = new LocalizationClass
			'	pointV.setLatitude(firstV(0))
			'	pointV.setLongitude(firstV(1))	
			'	objListVertices.add pointV, ""		
			'	Set pointV = nothing
			'end if    
			Set convertVertices=objListVertices
		end if
	end Function

	
	'****************************** START: FUNZIONE STATICA PER CONVERTIRE UN PUNTO (LAT,LON) IN UN OGGETTO LocalizationClass
	Public Function convertCenter(center)
		arrCenterPoint = Split(center, ",", -1, 1)
		'response.write("arrCenterPoint(0):"&arrCenterPoint(0)&" - arrCenterPoint(1):"&arrCenterPoint(1)&"<br>")
		Set pointCenterCircle = new LocalizationClass
		pointCenterCircle.setLatitude(arrCenterPoint(0))
		pointCenterCircle.setLongitude(arrCenterPoint(1))
		Set convertCenter = pointCenterCircle
		Set pointCenterCircle = nothing
	end Function


	'************************* START: FUNZIONI DI UTILITÀ TRIGONOMETRICHE
	' Find the great-circle distance in metres, assuming a spherical earth, between two lat-long points in degrees. */
	Private function greatCircleDistanceInMeters(aLong1, aLat1, aLong2, aLat2)
		KPiDouble = 3.14159265358979
		KDegreesToRadiansDouble = KPiDouble / 180.0
		'A constant to convert radians to metres for the Mercator and other projections.
		'It is the semi-major axis (equatorial radius) used by the WGS 84 datum (see http://en.wikipedia.org/wiki/WGS84).
		KEquatorialRadiusInMetres = 6378137

		aLong1 = aLong1*KDegreesToRadiansDouble
		aLat1 = aLat1*KDegreesToRadiansDouble
		aLong2 = aLong2*KDegreesToRadiansDouble
		aLat2 = aLat2*KDegreesToRadiansDouble

		angle = acos(sin(aLat1) * sin(aLat2) + cos(aLat1) * cos(aLat2) * cos(aLong2 - aLong1))

		greatCircleDistanceInMeters = angle * KEquatorialRadiusInMetres
	End function
	
	Private Function rad2deg(radians)	
		rad2deg = radians*180/pi
		'response.write("rad2deg:"&rad2deg&" -radians:"&radians&" -pi:"&pi&"<br>")
	End Function

	Private Function deg2rad(degrees)
		deg2rad = degrees*pi/180
		'response.write("deg2rad:"&deg2rad&" -degrees:"&degrees&" -pi:"&pi&"<br>")
	End Function

	Private Function pi()
		'pi=4*Atn(1)
		pi = 3.14159265358979
	end Function

	Private Function ATan2(y, x) 
		If x > 0 Then
			ATan2 = Atn(y / x)
		ElseIf x < 0 Then
			ATan2 = Sgn(y) * (pi - Atn(Abs(y / x)))
		ElseIf y = 0 Then
			ATan2 = 0
		Else
			ATan2 = Sgn(y) * pi / 2
		End If
	End Function

	' arc sine
	' error if value is outside the range [-1,1]
	Private Function ASin(value)
		If Abs(value) <> 1 Then
			ASin = Atn(value / Sqr(1 - value * value))
		Else
			ASin = 1.5707963267949 * Sgn(value)
		End If
	End Function

	' arc cosine
	' error if NUMBER is outside the range [-1,1]
	Private Function ACos(number)
		If Abs(number) <> 1 Then
			ACos = 1.5707963267949 - Atn(number / Sqr(1 - number * number))
		ElseIf number = -1 Then
			ACos = 3.14159265358979
		End If
		'elseif number=1 --> Acos=0 (implicit)
	End Function

	' arc cotangent
	' error if NUMBER is zero
	Private Function ACot(value) 
		ACot = Atn(1 / value)
	End Function

	' arc secant
	' error if value is inside the range [-1,1]
	Private Function ASec(value)
		' NOTE: the following lines can be replaced by a single call
		'            ASec = ACos(1 / value)
		If Abs(value) <> 1 Then
			ASec = 1.5707963267949 - Atn((1 / value) / Sqr(1 - 1 / (value * value)))
		Else
			ASec = 3.14159265358979 * Sgn(value)
		End If
	End Function

	' arc cosecant
	' error if value is inside the range [-1,1]
	Private Function ACsc(value)
		' NOTE: the following lines can be replaced by a single call
		'            ACsc = ASin(1 / value)
		If Abs(value) <> 1 Then
			ACsc = Atn((1 / value) / Sqr(1 - 1 / (value * value)))
		Else
			ACsc = 1.5707963267949 * Sgn(value)
		End If
	End Function
	'************************* END: FUNZIONI DI UTILITÀ TRIGONOMETRICHE
	
	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = doubleValue
		convertDoubleDelimiter = Replace(convertDoubleDelimiter, ".",",")		
	End Function

	Public Function revertDoubleDelimiter(doubleValue)
		revertDoubleDelimiter = doubleValue
		revertDoubleDelimiter = Replace(revertDoubleDelimiter, ",",".")		
	End Function
End Class
%>