<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->

<html>
  <head>
    <title>Google Maps JavaScript API v3 Example: Polygon Simple</title>
<META http-equiv="Content-Type" CONTENT="text/html; charset=utf-8">
<script src="https://maps.googleapis.com/maps/api/js?key=<%=Trim(Application("googlemaps_key"))%>&amp;sensor=false&libraries=drawing,geometry" type="text/javascript"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery-latest.min.js"%>"></script>
    <script>
    //<![CDATA[
    // delay between geocode requests - at the time of writing, 100 miliseconds seems to work well
    var delay = 100;

      var geo = new google.maps.Geocoder(); 

      // ====== Geocoding ======
      function getAddress(id_element, search, next) {
        geo.geocode({'address':search}, function (results,status)
          { 
            // If that was successful
            if (status == google.maps.GeocoderStatus.OK) {
              // Lets assume that the first marker is the one we want
              var p = results[0].geometry.location;
              var lat=p.lat();
              var lng=p.lng();
              // Output the data
		$('#inner').append("INSERT INTO googlemap_localization(id_element, `type`, latitude, longitude) VALUES("+id_element+",3,"+lat+","+lng+");<br>");
            }
            // ====== Decode the error status ======
            else {
              // === if we were sending the requests to fast, try this one again and increase the delay
              if (status == google.maps.GeocoderStatus.OVER_QUERY_LIMIT) {
                nextAddress--;
                delay++;
              } else {
                var reason=" - Code "+status;
		var msg = "error while parsing address: "+ search + " for element id: "+ id_element + reason;
		$('#inner').append(msg+"<br>");
              }   
            }
            next();
          }
        );
      }

      // ======= An array of locations that we want to Geocode ========
      var addresses = [
<%
Set objCountry = New CountryClass

On Error Resume Next
Set objListaCountry = objCountry.getListaCountry(null,null,null,null)				
counter = objListaCountry.count-1
for each x in objListaCountry

	if(objListaCountry(x).getStateRegionCode()<>"")then
		response.write("{key:'"&x&"',value:'"&Server.HTMLEncode(objListaCountry(x).getStateRegionDescription()) & " " & Server.HTMLEncode(objListaCountry(x).getCountryDescription())&"'}")
	else
		response.write("{key:'"&x&"',value:'"&Server.HTMLEncode(objListaCountry(x).getCountryDescription())&"'}")	
	end if
	if(counter>0)then
		response.write(",")
	end if
	counter = counter-1
next

Set objListaCountry = nothing
if Err.number <> 0 then
end if

Set objCountry = nothing
%>
      ];      
      

      // ======= Global variable to remind us what to do next
      var nextAddress = 0;

      // ======= Function to call the next Geocode operation when the reply comes back

      function theNext() {
        if (nextAddress < addresses.length) {
          setTimeout('getAddress("'+addresses[nextAddress].key+'","'+addresses[nextAddress].value+'",theNext)', delay);
          nextAddress++;
        }
      }

      // ======= Call that function for the first time =======
      theNext();

    // This Javascript is based on code provided by the
    // Community Church Javascript Team
    // http://www.bisphamchurch.org.uk/   
    // http://econym.org.uk/gmap/

    //]]>
    </script>
  </head>
  <body>
 <div id="inner">
 </div>
 
 
 <% 
Dim url, objHttp
url = "http://maps.googleapis.com/maps/api/geocode/xml?sensor=false&address=" 

On Error Resume Next
'**************************** CHIAMATA BASATA SU  XMLHTTP ************************ COMMENTATA NON  IN USO
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
'objHttp.open "GET", url, false
'objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
'objHttp.Send()
'Set objXML = objHTTP.ResponseXML
'response.write(objHTTP.responseText)
'set items = objXML.getElementsByTagName("location")
'lat = items(0).childNodes(0).text		
'lng = items(0).childNodes(1).text
'response.write("lat:"&lat&" -lng:"&lng)
'set items = nothing
'Set objXML = nothing
'set objHttp = nothing 


Set objXML = Server.CreateObject("Microsoft.XMLDOM")
objXML.setProperty "ServerHTTPRequest", true
objXML.async = False

Set objCountry = New CountryClass
Set objListaCountry = objCountry.getListaCountry(null,null,null,null)
ccounter=0
for each x in objListaCountry
	address = ""
	if(objListaCountry(x).getStateRegionCode()<>"")then
		address = Server.HTMLEncode(objListaCountry(x).getStateRegionDescription() & " " & objListaCountry(x).getCountryDescription())
		'address = Replace(address, " ", "+", 1, -1, 1)
		'address = Replace(address, "(", "", 1, -1, 1)
		'address = Replace(address, ")", "", 1, -1, 1)
	else
		address = Server.HTMLEncode(objListaCountry(x).getCountryDescription())
		'address = Replace(address, " ", "+", 1, -1, 1)
		'address = Replace(address, "(", "", 1, -1, 1)
		'address = Replace(address, ")", "", 1, -1, 1)
	end if	
	
	'objXML.Load (url&address)
	'lat = objXML.selectSingleNode("//location/lat").text
	'lng = objXML.selectSingleNode("//location/lng").text

	'response.write(address&"-----&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;INSERT INTO googlemap_localization(id_element, `type`, latitude, longitude) VALUES("&x&",3,"&lat&","&lng&");<br>")
	'response.write("INSERT INTO googlemap_localization(id_element, `type`, latitude, longitude) VALUES("&x&",3,"&lat&","&lng&");<br>")
	
	ccounter=ccounter+1
next
'response.write("<br><br><br>"&ccounter)
Set objListaCountry = nothing
Set objCountry = nothing

Set objXML = nothing

if(Err.number <> 0) then
response.write(Err.description)
end if 
%>
 
 
  </body>
</html>