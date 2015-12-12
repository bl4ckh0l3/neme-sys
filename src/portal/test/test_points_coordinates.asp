<!-- #include virtual="/common/include/Objects/LocalizationClass.asp" -->
<%
'verifico se un punto appartiene ad un poligono lato server
Public function pointInPolygon(point, Vertices)    
	j = Vertices.Count - 1  
	pointInPolygon = false  

	keys=Vertices.Keys

	for i=0 to Vertices.Count-1 
		if (keys(i).getLongitude() < point.getLongitude() AND keys(j).getLongitude() >= point.getLongitude() OR keys(j).getLongitude() < point.getLongitude() AND keys(i).getLongitude() >= point.getLongitude())  then  
			if (keys(i).getLatitude() +  (point.getLongitude() - keys(i).getLongitude())/(keys(j).getLongitude() - keys(i).getLongitude())*(keys(j).getLatitude() - keys(i).getLatitude()) < point.getLatitude())  then  
				pointInPolygon = not(pointInPolygon)  
			end if  
		end if  
		j = i 
	Next  
End function


'altra soluzione con due metodi
Private function isLeft(P0, P1, P2 )
    isLeft =  ( (P1.getLatitude() - P0.getLatitude()) * (P2.getLongitude() - P0.getLongitude()) - (P2.getLatitude() -  P0.getLatitude()) * (P1.getLongitude() - P0.getLongitude()) )
End function
Public function wn_PnPoly( point, Vertices)
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
	wn_PnPoly = false
    else
	wn_PnPoly = true
    end if
End function


'terza soluzione
Public function IsInside(punto, Vertices)
            x=Vertices.items
            y=Vertices.items
	    keys=Vertices.Keys

            px = punto.getLatitude()
            py = punto.getLongitude()
 
            crossings = 0
            
            for i=0 to Vertices.Count-1
                x(i) = keys(i).getLatitude()
                y(i) = keys(i).getLongitude()
            Next
            
            for i=0 to Vertices.Count-2
                if (x(i) < x((i + 1) mod (Vertices.Count - 1))) then
                        x1 = x(i)
                        x2 = x((i + 1) mod (Vertices.Count - 1))
                else 
                    x1 = x((i + 1) mod (Vertices.Count - 1))
		    x2 = x(i)
                end if
       
                if ( px > x1 AND px <= x2 AND ( py < y(i) OR py <= y((i+1)mod 8) ) ) then
                        eps =0.000001
 
                        dx = x((i + 1) mod (Vertices.Count - 1)) - x(i)
                        dy = y((i + 1) mod (Vertices.Count - 1)) - y(i)
 
                        if (Abs(dx) < eps) then
                            k = 3.402823e38
                        else
                            k = dy/dx
			end if
 
                        m = y(i) - k * x(i)
               
                        y2 = k * px + m
                        if ( py <= y2 ) then
                            crossings=crossings+1
			end if
                end if
            Next
 
            if (crossings mod 2 = 1) then
                IsInside= true
            else
                IsInside= false
	    end if
End function



' quarta soluzione NON FUNZIONA
Public function coordinate_is_inside_polygon(latitude, longitude, lat_array, long_array)
	angle=0
	int n = UBound(lat_array)

	for i=0 to n-1
		point1_lat = lat_array(i) - latitude
		point1_long = long_array(i) - longitude
		point2_lat = lat_array((i+1)mod n) - latitude
		point2_long = long_array((i+1)mod n) - longitude
		angle = angle+Angle2D(point1_lat,point1_long,point2_lat,point2_long)
	next

	if (Abs(angle) < PI) then
		coordinate_is_inside_polygon = false
	else
		coordinate_is_inside_polygon = true
	end if
End function

Private function Angle2D(y1, x1, y2, x2)
	Dim dtheta,theta1,theta2

	theta1 = atan2(y1,x1)
	theta2 = atan2(y2,x2)
	dtheta = theta2 - theta1
	Do While (dtheta > PI)
		dtheta = dtheta-(2*PI)
	Loop
	Do While (dtheta < -PI)
		dtheta = dtheta+(2*PI)
	Loop

	Angle2D = dtheta
End function



Private Function convertVertices(vertices)
  response.write("<br>vertices:"&vertices&"<br>")
	listVertices = Split(vertices, "|", -1, 1)	
	if(isArray(listVertices)) then
		Set objListVertices = Server.CreateObject("Scripting.Dictionary")
    firstV =""
		For y=LBound(listVertices) to UBound(listVertices)
     response.write("listVertices(y):"&listVertices(y)&"<br>")
			arrLatLon = Split(listVertices(y), ",", -1, 1)
      if(y=LBound(listVertices))then
        firstV = arrLatLon
      end if
			if(isArray(arrLatLon)) then
				Set pointV = new LocalizationClass
				pointV.setLatitude(arrLatLon(0))
				pointV.setLongitude(arrLatLon(1))	
				objListVertices.add pointV, ""		
        Set pointV = nothing
			end if
		next
    if(isArray(firstV)) then
      Set pointV = new LocalizationClass
      pointV.setLatitude(firstV(0))
      pointV.setLongitude(firstV(1))	
      objListVertices.add pointV, ""		
      Set pointV = nothing
    end if    
		Set convertVertices=objListVertices
	end if
end Function

Private Function convertVerticesToArr(vertices)
  response.write("<br>vertices:"&vertices&"<br>")
	listVertices = Split(vertices, "|", -1, 1)	
	if(isArray(listVertices)) then
		Set objListVertices = Server.CreateObject("Scripting.Dictionary")
		For y=LBound(listVertices) to UBound(listVertices)
      response.write("listVertices(y):"&listVertices(y)&"<br>")
			arrLatLon = Split(listVertices(y), ",", -1, 1)

			if(isArray(arrLatLon)) then
				objListVertices.add arrLatLon(0), arrLatLon(1)		
			end if
		next
		Set convertVertices=objListVertices
	end if
end Function

Private Function convertCenter(center)
	arrCenterPoint = Split(center, ",", -1, 1)
  'response.write("arrCenterPoint(0):"&arrCenterPoint(0)&" - arrCenterPoint(1):"&arrCenterPoint(1)&"<br>")
	Set pointCenterCircle = new LocalizationClass
	pointCenterCircle.setLatitude(arrCenterPoint(0))
	pointCenterCircle.setLongitude(arrCenterPoint(1))

  Set convertCenter = pointCenterCircle
  Set pointCenterCircle = nothing
end Function







'***************************************************************** FUNZIONI DI VERIFICA APPARTENENZA PUNTO AD UN CERCHIO *********************************

'************************* START: FUNZIONE PER DETERMINARE SE UN PUNTO APPARTIENE AD UNA CIRCONFERENZA SULLA SUPERFICIE TERRESTRE
'************************* NON FUNZIONA
' square-root((x1-xc)^2 + (y1-yc)^2)) < R
' oppure
' (x1-xc)^2 + (y1-yc)^2) < R^2
Public function IsInsideCircle(punto, center, radius)

	deltaX = punto.getLatitude()-center.getLatitude()
	deltaY = punto.getLongitude()-center.getLongitude()

	'response.write("<br><br>deltaX:"&deltaX&"<br>")
	'response.write("deltaY:"&deltaY&"<br>")

	'deltaX2 =deltaX^2
	'deltaY2 =deltaY^2 
	deltaX2 =deltaX*deltaX
	deltaY2 =deltaY*deltaY 

	'response.write("<br>deltaX2:"&deltaX2&"<br>")
	'response.write("deltaY2:"&deltaY2&"<br>")

	deltaSum = deltaX2+deltaY2
	'response.write("deltaSum:"&deltaSum&"<br>")
	'response.write("radius:"&radius&"<br>")
	'response.write("Sqr(deltaSum):"&Sqr(deltaSum)&"<br>")

	if (deltaSum <= radius*radius) then
		IsInsideCircle=true
	else
		IsInsideCircle=false
	end if
end Function
'************************* END: FUNZIONE PER DETERMINARE SE UN PUNTO APPARTIENE AD UNA CIRCONFERENZA SULLA SUPERFICIE TERRESTRE
'************************* NON FUNZIONA


'************************* START: ALTRA FUNZIONE PER DETERMINARE SE UN PUNTO APPARTIENE AD UNA CIRCONFERENZA SULLA SUPERFICIE TERRESTRE
'************************* NON FUNZIONA
Public function IsInsideCircle2(punto, center, radius)
	arrB = getBoundingBox(center, radius)

	bolX = (punto.getLatitude() >= arrB(0) AND punto.getLatitude() <= arrB(1))
	bolY = (punto.getLongitude()  >= arrB(2) AND punto.getLongitude() <= arrB(3))

	if (bolX AND bolY) then
		IsInsideCircle2=true
	else
		IsInsideCircle2=false
	end if
end Function
' given a latitude and longitude in degrees (40.123123,-72.234234) and a distance in miles
' calculates a bounding box with corners $distance_in_miles away from the point specified.
' returns $min_lat,$max_lat,$min_lon,$max_lon 
Private function getBoundingBox(center, radiusParam)
	Dim arrbox(4)
	'earth_radius = 6377991.2064 'of earth in meters
	earth_radius = 3963.1 ' of earth in miles
	
	radiusParam = radiusParam * 0.00062137 ' radius in miles

	' bearings	
	due_north = 0
	due_south = 180
	due_east = 90
	due_west = 270

	' convert latitude and longitude into radians 
	lat_r = deg2rad(center.getLatitude())
	lon_r = deg2rad(center.getLongitude())
		
	' find the northmost, southmost, eastmost and westmost corners $distance_in_miles away
	' original formula from
	' http://www.movable-type.co.uk/scripts/latlong.html

	northmost  = asin(sin(lat_r) * cos(radiusParam/earth_radius) + cos(lat_r) * sin (radiusParam/earth_radius) * cos(due_north))
	southmost  = asin(sin(lat_r) * cos(radiusParam/earth_radius) + cos(lat_r) * sin (radiusParam/earth_radius) * cos(due_south))
	
	eastmost = lon_r + atan2(sin(due_east)*sin(radiusParam/earth_radius)*cos(lat_r),cos(radiusParam/earth_radius)-sin(lat_r)*sin(lat_r))
	westmost = lon_r + atan2(sin(due_west)*sin(radiusParam/earth_radius)*cos(lat_r),cos(radiusParam/earth_radius)-sin(lat_r)*sin(lat_r))
		
		
	northmost = rad2deg(northmost)
	southmost = rad2deg(southmost)
	eastmost = rad2deg(eastmost)
	westmost = rad2deg(westmost)
		
	' sort the lat and long so that we can use them for a between query		
	if (northmost > southmost) then 
		lat1 = southmost
		lat2 = northmost
	
	 else
		lat1 = northmost
		lat2 = southmost
	end if


	if (eastmost > westmost) then 
		lon1 = westmost
		lon2 = eastmost
	
	else
		lon1 = eastmost
		lon2 = westmost
	end if
	
	arrbox(0) = lat1
	arrbox(1) = lat2
	arrbox(2) = lon1
	arrbox(3) = lon2
	getBoundingBox =  arrbox
End function
'************************* END: ALTRA FUNZIONE PER DETERMINARE SE UN PUNTO APPARTIENE AD UNA CIRCONFERENZA SULLA SUPERFICIE TERRESTRE
'************************* NON FUNZIONA


'************************* START: FUNZIONI DI UTILITÀ TRIGONOMETRICHE
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


Set objListaPoint = Server.CreateObject("Scripting.Dictionary")

Set objLoc1 = new LocalizationClass	
objLoc1.setLatitude(25.774252)
objLoc1.setLongitude(-80.190262)	
Set objLoc2 = new LocalizationClass	
objLoc2.setLatitude(18.466465)
objLoc2.setLongitude(-66.118292)	
Set objLoc3 = new LocalizationClass	
objLoc3.setLatitude(32.321384)
objLoc3.setLongitude(-64.75737)	
Set objLoc4 = new LocalizationClass	
objLoc4.setLatitude(25.774252)
objLoc4.setLongitude(-80.190262)	
objListaPoint.add objLoc1, ""	
objListaPoint.add objLoc2, ""	
objListaPoint.add objLoc3, ""
objListaPoint.add objLoc4, ""

Set pointCheck = new LocalizationClass
pointCheck.setLatitude(26.980829)
pointCheck.setLongitude(-70.052491)

Set pointCheck2 = new LocalizationClass
pointCheck2.setLatitude(25.774252)
pointCheck2.setLongitude(-80.190262)

Set pointCheck3 = new LocalizationClass
pointCheck3.setLatitude(32.268555)
pointCheck3.setLongitude(-59.088135)

Set pointCheck4 = new LocalizationClass
pointCheck4.setLatitude(29.190533)
pointCheck4.setLongitude(-69.920655)


'************************************** check on circle
Set pointCenterCircle = new LocalizationClass
pointCenterCircle.setLatitude(25.774252)
pointCenterCircle.setLongitude(-80.190262)

Set pointCheckCircle = new LocalizationClass
pointCheckCircle.setLatitude(32.268555)
pointCheckCircle.setLongitude(-59.088135)

Set pointCheckCircle2 = new LocalizationClass
pointCheckCircle2.setLatitude(25.799891)
pointCheckCircle2.setLongitude(-80.291748)

Set pointCheckCircle3 = new LocalizationClass
pointCheckCircle3.setLatitude(18.466465)
pointCheckCircle3.setLongitude(-66.118292)

radius = 2000000 ' 2000 km
%>


<html>
  <head>
    <title>Google Maps JavaScript API v3 Example: Polygon Simple</title>
<script src="https://maps.googleapis.com/maps/api/js?key=<%=Trim(Application("googlemaps_key"))%>&amp;sensor=false&libraries=drawing,geometry" type="text/javascript"></script>
    <script>
var infoWindow;
var map;

      function initialize() {
        var myLatLng = new google.maps.LatLng(0, 0);
        var mapOptions = {
          zoom: 2,
          center: myLatLng,
          mapTypeId: google.maps.MapTypeId.ROADMAP
        };
        map = new google.maps.Map(document.getElementById('map_canvas'), mapOptions);

	infowindow = new google.maps.InfoWindow();
	
        var bermudaTriangle;
        var triangleCoords = [
            new google.maps.LatLng(25.774252, -80.190262),
            new google.maps.LatLng(18.466465, -66.118292),
            new google.maps.LatLng(32.321384, -64.75737),
            new google.maps.LatLng(25.774252, -80.190262)
        ];

        // Construct the polygon
        bermudaTriangle = new google.maps.Polygon({
          paths: triangleCoords,
          strokeColor: '#FF0000',
          strokeOpacity: 0.8,
          strokeWeight: 2,
          fillColor: '#FF0000',
          fillOpacity: 0.35
        });

        bermudaTriangle.setMap(map);

	var testcontains = google.maps.geometry.poly.containsLocation(new google.maps.LatLng(26.980829, -70.052491), bermudaTriangle);
	//document.write("js - contains1: "+testcontains+"<br>");
	testcontains = google.maps.geometry.poly.containsLocation(new google.maps.LatLng(25.774252, -80.190262), bermudaTriangle);
	//document.write("js - contains2: "+testcontains+"<br>");
	testcontains = google.maps.geometry.poly.containsLocation(new google.maps.LatLng(32.268555,-59.088135), bermudaTriangle);
	//document.write("js - contains3: "+testcontains+"<br>");
	

	/*var drawingManager = new google.maps.drawing.DrawingManager({
		drawingMode: google.maps.drawing.OverlayType.POLYGON,
		drawingControl: true,
		drawingControlOptions: {
			position: google.maps.ControlPosition.TOP_CENTER,
			drawingModes: [
				//google.maps.drawing.OverlayType.MARKER,
				//google.maps.drawing.OverlayType.CIRCLE,
				google.maps.drawing.OverlayType.POLYGON,
				//google.maps.drawing.OverlayType.POLYLINE,
				//google.maps.drawing.OverlayType.RECTANGLE
			]
		},
		circleOptions: {
			fillColor: '#ffff00',
			fillOpacity: 1,
			strokeWeight: 5,
			clickable: false,
			zIndex: 1,
			editable: true
		},
		polygonOptions: {
			editable: true			
		}
	});

	drawingManager.setMap(map);

	google.maps.event.addListener(drawingManager, 'overlaycomplete', function(event) {
	if (event.type == google.maps.drawing.OverlayType.CIRCLE) {
		var radius = event.overlay.getRadius();
	}
	else if (event.type == google.maps.drawing.OverlayType.POLYGON) {
		showArrays(event);
		//drawingManager.setOptions({
		//	drawingControl: false,
		//	drawingMode: null
		//});
	}
	});*/




	// Construct the circle for each value in citymap. We scale population by 20.
	var populationOptions = {
	strokeColor: "#FF0000",
	strokeOpacity: 0.8,
	strokeWeight: 2,
	fillColor: "#FF0000",
	fillOpacity: 0.35,
	map: map,
	center: new google.maps.LatLng(25.774252, -80.190262),
	radius: 2000000 //2000 Km
	};
	cityCircle = new google.maps.Circle(populationOptions);

createMarker(new google.maps.LatLng(32.268555,-59.088135),map);
createMarker(new google.maps.LatLng(25.799891,-80.291748),map);
createMarker(new google.maps.LatLng(18.466465,-66.118292),map);

	bounds = cityCircle.getBounds();
	testcontains = bounds.contains(new google.maps.LatLng(32.268555,-59.088135));
	//alert("contains circle: "+testcontains);
	testcontains = bounds.contains(new google.maps.LatLng(25.799891,-80.291748));
	//alert("contains circle 2: "+testcontains);
	testcontains = bounds.contains(new google.maps.LatLng(18.466465,-66.118292));
	//alert("contains circle 3: "+testcontains);


	//disegno un cerchio e cambio il raggio , dovrebbe avere un raggio su base kilometrica 
	//play(new google.maps.LatLng(45.836454,9.314575), 1000);
	//play(new google.maps.LatLng(32.268555,-59.088135), 50);


}
 
// selector function
function play(center, radius){
	var circleOptions = {
	strokeColor: "#FF0000",
	strokeOpacity: 0.8,
	strokeWeight: 2,
	fillColor: "#FF0000",
	fillOpacity: 0.35,
	map: map,
	center: center,
	radius: radius * 1000
	};
	circle = new google.maps.Circle(circleOptions);
	
	map.fitBounds(circle.getBounds());
	map.circleRadius = radius;
}




function showArrays(event) {

  // Since this Polygon only has one path, we can call getPath()
  // to return the MVCArray of LatLngs
  var vertices = event.overlay.getPath();

  var contentString = "<b>Polygon</b><br />";
  //contentString += "Clicked Location: <br />" + event.overlay.latLng.lat() + "," + event.overlay.latLng.lng() + "<br />";

  // Iterate over the vertices.
  for (var i =0; i < vertices.length; i++) {
    var xy = vertices.getAt(i);
    contentString += "<br />" + "Coordinate: " + i + "<br />" + xy.lat() +"," + xy.lng();
  }

  // Replace our Info Window's content and position
  infowindow.setContent(contentString);
  infowindow.setPosition(vertices.getAt(0));

  infowindow.open(map);
}
 
function createMarker(point,map) {  
	var infowindow = new google.maps.InfoWindow(); 
	var marker = new google.maps.Marker({
		position: point,
		map: map
	});
	return marker;
} 
    </script>
  </head>
  <body onload="initialize()">
    <div style="width:550px;height:500px;vertical-align:top;text-align:left;border:1px solid;background:#FFFFFF;" id="map_canvas"></div>
    <br><br><br>
 <%
 

'response.write("function 1 - contains1: "& pointInPolygon(pointCheck, objListaPoint) &"<br>")
'response.write("function 1 - contains2: "& pointInPolygon(pointCheck2, objListaPoint) &"<br>")
'response.write("function 1 - contains3: "& pointInPolygon(pointCheck3, objListaPoint) &"<br>")
'response.write("function 1 - contains4: "& pointInPolygon(pointCheck4, objListaPoint) &"<br>")

'response.write("function 2 - contains1: "& wn_PnPoly(pointCheck, objListaPoint) &"<br>")
'response.write("function 2 - contains2: "& wn_PnPoly(pointCheck2, objListaPoint) &"<br>")
'response.write("function 2 - contains3: "& wn_PnPoly(pointCheck3, objListaPoint) &"<br>")
'response.write("function 2 - contains4: "& wn_PnPoly(pointCheck4, objListaPoint) &"<br>")

'response.write("function 3 - contains1: "& IsInside(pointCheck, objListaPoint) &"<br>")
'response.write("function 3 - contains2: "& IsInside(pointCheck2, objListaPoint) &"<br>")
'response.write("function 3 - contains3: "& IsInside(pointCheck3, objListaPoint) &"<br>")
'response.write("function 3 - contains4: "& IsInside(pointCheck4, objListaPoint) &"<br>")


'response.write("function circle - contains1: "& IsInsideCircle(pointCheckCircle, pointCenterCircle, radius) &"<br>")
'response.write("function circle - contains2: "& IsInsideCircle(pointCheckCircle2, pointCenterCircle, radius) &"<br>")
'response.write("function circle - contains3: "& IsInsideCircle(pointCheckCircle3, pointCenterCircle, radius) &"<br>")

Set objLocBase = new LocalizationClass

response.write("objLocBase.isPointInPolygon(): "& objLocBase.isPointInPolygon(pointCheck, objListaPoint) &"<br>")
response.write("objLocBase.isPointInPolygon(): "& objLocBase.isPointInPolygon(pointCheck2, objListaPoint) &"<br>")
response.write("objLocBase.isPointInPolygon(): "& objLocBase.isPointInPolygon(pointCheck3, objListaPoint) &"<br>")
response.write("objLocBase.isPointInPolygon(): "& objLocBase.isPointInPolygon(pointCheck4, objListaPoint) &"<br>")

response.write("objLocBase.IsPointInCircle(): "& objLocBase.isPointInCircleOnEarthSurface(pointCheckCircle, pointCenterCircle, radius) &"<br>")
response.write("objLocBase.IsPointInCircle(): "& objLocBase.isPointInCircleOnEarthSurface(pointCheckCircle2, pointCenterCircle, radius) &"<br>")
response.write("objLocBase.IsPointInCircle(): "& objLocBase.isPointInCircleOnEarthSurface(pointCheckCircle3, pointCenterCircle, radius) &"<br>")

Set objLocBase = nothing



'test of decimal separators
a=45.696588248373764
b= 8.983383178710938
c="8,54"
response.write("<br> typename(c):"& typename(c)&" -value:"&c)
c=Cdbl(c)
response.write("<br> typename(c):"& typename(c)&" -value:"&c)

response.write("<br> a maggiore b:"& (a>b))


response.write("<br> substract a-b:"& a-b)

response.write("<br> multiply:"& 4335 * 90 / 10000000)
 %>   
  </body>
</html>