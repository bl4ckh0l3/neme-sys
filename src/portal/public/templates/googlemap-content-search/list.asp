<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->
<!-- #include file="include/init1.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=pageTemplateTitle%></title>
<META name="description" CONTENT="<%=metaDescription%>">
<META name="keywords" CONTENT="<%=metaKeyword%>">
<META name="autore" CONTENT="Neme-sys; email:info@neme-sys.org">
<META http-equiv="Content-Type" CONTENT="text/html; charset=utf-8">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<%if not(isNull(strCSS)) ANd not(strCSS = "") then%><link rel="stylesheet" href="<%=Application("baseroot") & strCSS%>" type="text/css"><%end if%>
<script language="Javascript">
var mapid = "maplist";
var latlng = new Array();
var infowin = new Array();
var markers = new Array();
var map, mc;
var drawingManager;
var lastSelectionType="";
var currentOverlay = new Array();
var currentOverlayCoordinates;
var hasGeoSearchActive = false;
<%if(objListPoint.count>0)then
	for each k in objListPoint
		//response.write("//getLatitude: "&k.getLatitude()&" -getLongitude: "&k.getLongitude())
		if(k.getLatitude()<>"" AND k.getLongitude()<>"")then%>
			latlng.push(new google.maps.LatLng(replaceCommaInNumber('<%=k.getLatitude()%>'), replaceCommaInNumber('<%=k.getLongitude()%>')));
			infowin.push("<%=objListPoint(k)%>");
		<%end if
	next%>
<%end if%>

function showMap(mapid){
	$('#'+mapid).show();
        var mapOptions = {
          center: new google.maps.LatLng(41.87194,12.567379999999957),
          zoom: 5,
          mapTypeId: google.maps.MapTypeId.ROADMAP
        };
        map = new google.maps.Map(document.getElementById(mapid),  mapOptions);
	var latlngbounds = new google.maps.LatLngBounds();

	if(latlng.length==1){
		map.setCenter(latlng[0]);
		map.setZoom(10);
	}else if(latlng.length>1){
		for (var j=0; j<latlng.length; j++){
			latlngbounds.extend(latlng[j]);
		}
		map.fitBounds(latlngbounds);
	}

	for (var j=0; j<latlng.length; j++){
		var infowintxt = infowin[j];
		var marker = createMarker(latlng[j], infowintxt, map);
	}
	
	setDrawingManager(map);
	
	<%
	if (bolHasGeoSearchActive)then
		response.write("createShapes('"&Session("geolocalsearchpoly")("current_overlay")&"', map, true);")
		response.write("lastSelectionType='"&Session("geolocalsearchpoly")("last_selection")&"';")
	end if
	%>
	
	var mcOptions = {gridSize: 50, maxZoom: 15};
	mc = new MarkerClusterer(map, markers, mcOptions);
}


function setDrawingManager(mapObj){
	drawingManager = new google.maps.drawing.DrawingManager({
		drawingMode: google.maps.drawing.OverlayType.POLYGON,
		drawingControl: true,
		drawingControlOptions: {
			position: google.maps.ControlPosition.TOP_CENTER,
			drawingModes: [
				//google.maps.drawing.OverlayType.MARKER,
				google.maps.drawing.OverlayType.CIRCLE,
				google.maps.drawing.OverlayType.POLYGON,
				//google.maps.drawing.OverlayType.POLYLINE,
				//google.maps.drawing.OverlayType.RECTANGLE
			]
		},
		circleOptions: {
			/*fillColor: '#ffff00',
			fillOpacity: 1,
			strokeWeight: 5,*/
			clickable: false,
			zIndex: 1,
			editable: true
		},
		polygonOptions: {
			zIndex: 1,
			editable: true			
		}
	});

	drawingManager.setMap(mapObj);

	google.maps.event.addListener(drawingManager, 'overlaycomplete', function(event) {
		if (event.type == google.maps.drawing.OverlayType.CIRCLE) {
			setCircle(event);
			lastSelectionType = event.type;
			currentOverlay.push(event.overlay);
			drawingManager.setOptions({
				drawingControl: false,
				drawingMode: null
			});
			//$('#geosearchbuttons').show();
			$('#georesetbuttons').show();
			hasGeoSearchActive=true;	
		}
		else if (event.type == google.maps.drawing.OverlayType.POLYGON) {
			setVertices(event);
			lastSelectionType = event.type;
			currentOverlay.push(event.overlay);
			drawingManager.setOptions({
				drawingControl: false,
				drawingMode: null
			});
			//$('#geosearchbuttons').show();
			$('#georesetbuttons').show();	
			hasGeoSearchActive=true;
		}
	});
}

function setVertices(event) {
	var type = 1;
	var vertices = event.overlay.getPath();
	currentOverlayCoordinates=type+"#";
	var contentString = "type="+type+"&vertices=";
	for (var i =0; i < vertices.length; i++) {
		var xy = vertices.getAt(i);
		contentString += xy.lat() +"," + xy.lng() + "|";		
		currentOverlayCoordinates+= xy.lat() +"," + xy.lng() + "|";
		/*if(i < vertices.length-1){
			contentString +="|";
			currentOverlayCoordinates +="|";
		}*/
	}
	contentString += vertices.getAt(0).lat() +"," + vertices.getAt(0).lng();
	currentOverlayCoordinates+= vertices.getAt(0).lat() +"," + vertices.getAt(0).lng();
	
	dataString = contentString+"&current_overlay="+currentOverlayCoordinates+"&last_selection="+event.type; 
	//alert(dataString);	
	$.ajax({  
		type: "POST",  
		url: "<%=tmpurl&"ajaxsetgeolocalsearch.asp"%>",  
		data: dataString,  
		success: function(response) {  
			//$('#'+container).html(response); 
			//alert("funziona");
			document.form_geo_search.search_active.value=1;
		}
	}); 
}

function setCircle(event){
	var type = 2;
	var radius = parseInt(event.overlay.getRadius());
	var center = event.overlay.getCenter();
	currentOverlayCoordinates=type+"#"+radius+"#"+center.lat()+","+center.lng();
	var contentString = "";
	contentString+="type="+type+"&radius="+radius+"&center="+center.lat()+","+center.lng();
	dataString = contentString+"&current_overlay="+currentOverlayCoordinates+"&last_selection="+event.type;   
	//alert(dataString);
	$.ajax({  
		type: "POST",  
		url: "<%=tmpurl&"ajaxsetgeolocalsearch.asp"%>",  
		data: dataString,  
		success: function(response) {  
			//$('#'+container).html(response); 
			//alert("funziona");
			document.form_geo_search.search_active.value=1;
		}
	}); 	
}

function createMarker(point,html,map) {  
	var infowindow = new google.maps.InfoWindow(); 
	var marker = new google.maps.Marker({
		position: point,
		map: map
	});
	google.maps.event.addListener(marker, "click", function() {
		infowindow.setContent(html);
		infowindow.open(map, marker);					
	});
	markers.push(marker);
	return marker;
}

function reActivateDrawingMode(){
	for(var i=0; i<currentOverlay.length;i++){
		currentOverlay[i].setMap(null);
	}
	for(var i=0; i<markers.length;i++){
		markers[i].setMap(null);
	}
	mc.removeMarkers(markers);
	
	currentOverlayCoordinates="";	
	dataString = "type=0"; 
	$('#searchresetbuttons').hide();
	$('#georesetbuttons').hide();
	document.form_geo_search.search_active.value=0;
	$('#content_list_container').hide();
	hasGeoSearchActive=false;
	$.ajax({  
		type: "POST",  
		url: "<%=tmpurl&"ajaxsetgeolocalsearch.asp"%>",  
		data: dataString,  
		success: function(response) { 
				if($('#contentfield__<%=idrpv%>').val()!= ""){
				drawingManager.setOptions({
					drawingControl: true,
					drawingMode: lastSelectionType
				});
			}
		}
	}); 
}

function enableDrawingMode(){
	if(drawingManager){
		var selectionMode = lastSelectionType;
		if(selectionMode==""){
			selectionMode = google.maps.drawing.OverlayType.POLYGON;
		}
		drawingManager.setOptions({
			drawingControl: true,
			drawingMode: selectionMode
		});		
	}
}

function disableDrawingMode(){
	if(drawingManager){
		drawingManager.setOptions({
			drawingControl: false,
			drawingMode: null
		});		
	}
}

function sendGeoSearch(){
	if(document.form_geo_search.price_from.value!= "" || document.form_geo_search.price_to.value !=""){
		document.form_geo_search.field_price__<%=idprc%>.value = document.form_geo_search.price_from.value + ' x ' + document.form_geo_search.price_to.value;
	}else{
		document.form_geo_search.field_price__<%=idprc%>.value ="";
	}
	
	if(document.form_geo_search.superficie_from.value!= "" || document.form_geo_search.superficie_to.value !=""){
		document.form_geo_search.field_superficie__<%=idsup%>.value = document.form_geo_search.superficie_from.value + ' x ' + document.form_geo_search.superficie_to.value;
	}else{
		document.form_geo_search.field_superficie__<%=idsup%>.value ="";
	}

	if(document.form_geo_search.locali_from.value!= "" || document.form_geo_search.locali_to.value !=""){
		document.form_geo_search.field_locali__<%=idloc%>.value = document.form_geo_search.locali_from.value + ' x ' + document.form_geo_search.locali_to.value;
	}else{
		document.form_geo_search.field_locali__<%=idloc%>.value ="";
	}

	form_geo_search.submit();
}

function resetGeoSearch(){
	for(var i=0; i<currentOverlay.length;i++){
		currentOverlay[i].setMap(null);
	}
	currentOverlayCoordinates="";	
	dataString = "type=3"; 
	$('#georesetbuttons').hide();
	document.form_geo_search.search_active.value=0;
	hasGeoSearchActive=false;
	$.ajax({  
		type: "POST",  
		url: "<%=tmpurl&"ajaxsetgeolocalsearch.asp"%>",  
		data: dataString,  
		success: function(response) {  
			drawingManager.setOptions({
				drawingControl: true,
				drawingMode: lastSelectionType
			});		
		}
	}); 	
}

function createShapes(referStr, map, exist){
	var arr=referStr.split("#");
	
	if(arr[0]=="1"){
		var arrVertices=arr[1].split("|");
		createPolygon(arrVertices, map, exist);
	}else if(arr[0]=="2"){
		var radius = arr[1];
		var center = arr[2].split(",");
		createCircle(radius, center, map, exist);
	}
	
	if(exist){
		hasGeoSearchActive = true;
	}
}

function createPolygon(vertices, map, exist){
	var polyCoords = [];
	var latlngbounds = new google.maps.LatLngBounds();
	
	for(var i=0; i<vertices.length;i++){
		coords = vertices[i].split(",");
		var point = new google.maps.LatLng(coords[0], coords[1]);
		polyCoords.push(point);
		latlngbounds.extend(point);
	}

	newPoly = new google.maps.Polygon({
		paths: polyCoords/*,
		strokeColor: '#FF0000',
		strokeOpacity: 0.8,
		strokeWeight: 2,
		fillColor: '#FF0000',
		fillOpacity: 0.35*/
	});

	newPoly.setMap(map);
	map.fitBounds(latlngbounds);
	currentOverlay.push(newPoly);

	if(exist){
		drawingManager.setOptions({
			drawingControl: false,
			drawingMode: null
		});
		$('#georesetbuttons').show();		
	}
}

function createCircle(radiusPar, centerArr, map, exist){
	var circleOptions = {
	/*strokeColor: "#FF0000",
	strokeOpacity: 0.8,
	strokeWeight: 2,
	fillColor: "#FF0000",
	fillOpacity: 0.35,*/
	map: map,
	center: new google.maps.LatLng(centerArr[0],centerArr[1]),
	radius: parseInt(radiusPar)
	};
	newCircle = new google.maps.Circle(circleOptions);
	map.fitBounds(newCircle.getBounds());
	currentOverlay.push(newCircle);

	if(exist){
		drawingManager.setOptions({
			drawingControl: false,
			drawingMode: null
		});
		$('#georesetbuttons').show();
	}

}

function replaceCommaInNumber(number){
	return number.replace(',','.');
}
  
function openDetailContentPage(strAction, strGerarchia, numIdNews, numPageNum){
    document.form_detail_link_news.action=strAction;
    document.form_detail_link_news.gerarchia.value=strGerarchia;
    document.form_detail_link_news.id_news.value=numIdNews;
    document.form_detail_link_news.modelPageNum.value=numPageNum;
    document.form_detail_link_news.submit();
}

function ajaxLoadFilter(id_filter, description, target_lang, sorting, url_param){
	//var query_string = "field_desc="+field_desc+"&target_lang="+encodeURIComponent(field_val)+"&target_lang="+target_lang;
	if(url_param!=""){
		var split_url_param=url_param.split("&");
		url_param=""
		for (var i = 0; i < split_url_param.length; i++) {
			var keyval= split_url_param[i].split("=");
			//alert(split_url_param[i]);
			url_param+="&"+keyval[0]+"="+encodeURIComponent(keyval[1]);
		}		
	}
	
	var query_string = "description="+description+"&sorting="+sorting+"&target_lang="+target_lang+url_param;
	//alert("id_filter:"+id_filter+" -query_string: "+query_string);

	$.ajax({
		async: true,
		type: "POST",
		cache: false,
		url: "<%=Application("baseroot") &tmpurl&"ajaxloadfilter.asp"%>",
		data: query_string,
		success: function(response) {
			//alert("response: "+response);
			//$("#ajaxresp").empty();
			$("#"+id_filter).append(response);
			//$("#ajaxresp").fadeIn(1500,"linear");
			//$("#ajaxresp").fadeOut(600,"linear");
		},
		error: function() {
			//alert("error");
		}
	});
}

function ajaxLoadLatLng(mapObj, idEl,zoomLevel){
	var point = new google.maps.LatLng(41.87194,12.567379999999957);

	if(idEl!=""){
		var query_string = "id="+idEl;	
		$.ajax({
			async: false,
			type: "POST",
			cache: false,
			url: "<%=Application("baseroot") &tmpurl&"ajaxloadjselement.asp"%>",
			data: query_string,
			success: function(response) {
				var newarrpoint = eval(response);

				if(newarrpoint.length==1){
					mapObj.setCenter(newarrpoint[0]);
					mapObj.setZoom(zoomLevel);
				}else if(newarrpoint.length>1){			
					var latlngbounds = new google.maps.LatLngBounds();
					for (var j=0; j<newarrpoint.length; j++){
						latlngbounds.extend(newarrpoint[j]);
					}
					mapObj.fitBounds(latlngbounds);
				}			
			},
			error: function() {
				mapObj.setCenter(point);
				mapObj.setZoom(5);
			}
		});
	}else{
		mapObj.setCenter(point);
		mapObj.setZoom(5);		
	}
}

jQuery(document).ready(function(){
	<%if not(bolHasFilterSearchActive)then%>
	$('#searchresetbuttons').hide();
	<%end if%>
	$('#georesetbuttons').hide();

	showMap(mapid);
		
	if($('#contentfield__<%=idrpv%>').val()== ""){		
		disableDrawingMode();
	}
});
</script> 
</head>
<body>
<div id="warp">
	<!-- #include virtual="/public/layout/include/header.inc" -->	
	<div id="container">	
		<!-- include virtual="/public/layout/include/menu_orizz.inc" -->
		<!-- #include virtual="/public/layout/include/menu_vert_sx.inc" -->
		<div id="content-center">
			<!-- #include virtual="/public/layout/include/menutips.inc" -->

			<div align="left" id="contenuti">
				<br/>
				<div id="content-center-prodotto">	
					<div>

					<form action="<%=tmpurl&"list.asp"%>" method="post" name="form_geo_search">	
						<input type="hidden" value="" name="modelPageNum">	
						<input type="hidden" value="" name="gerarchia">	
						<input type="hidden" value="1" name="page">
						<input type="hidden" value="<%=order_by%>" name="order_by">				
						<input type="hidden" value="0" name="search_active">				
						<input type="hidden" value="1" name="fields_filter">

						<div id="contract" style="float:left;">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.contract")%></span><br/>
							<%							
							fieldValueMatch = objListPairKeyValue(CStr(idcon))
							response.write(objContentF.renderContentFieldHTML(objCON,"", "__", "", fieldValueMatch,lang,1,objCON.getEditable()))								
							%>
						</div>

						<div id="category" style="float:left;padding-left:20px;">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.category")%></span><br/>
							<%							
							fieldValueMatch = objListPairKeyValue(CStr(idcat))
							response.write(objContentF.renderContentFieldHTML(objCAT,"", "__", "", fieldValueMatch,lang,1,objCAT.getEditable()))								
							%>
						</div>

						<div id="typology" style="padding-left:20px;float:left;height:80px;">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.typology")%></span><br/>
							<%							
							fieldValueMatch = objListPairKeyValue(CStr(idtyp))
							response.write(objContentF.renderContentFieldHTML(objTYP,"", "__", "", fieldValueMatch,lang,1,objTYP.getEditable()))								
							%>
						</div>
						
						<div class="clear" style="height:10px;"></div>
						
						<div id="typeproperty" style="float:left;">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.typeproperty")%></span><br/>
							<%							
							fieldValueMatch = objListPairKeyValue(CStr(idtpp))
							response.write(objContentF.renderContentFieldHTML(objTPP,"", "__", "", fieldValueMatch,lang,1,objTPP.getEditable()))								
							%>
						</div>

						<div id="status" style="float:left;padding-left:20px;">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.status")%></span><br/>
							<%							
							fieldValueMatch = objListPairKeyValue(CStr(idsta))
							response.write(objContentF.renderContentFieldHTML(objSTA,"", "__", "", fieldValueMatch,lang,1,objSTA.getEditable()))								
							%>
						</div>

						<div id="riscaldamento" style="float:left;padding-left:20px;">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.riscaldamento")%></span><br/>
							<%							
							fieldValueMatch = objListPairKeyValue(CStr(idris))
							response.write(objContentF.renderContentFieldHTML(objRIS,"", "__", "", fieldValueMatch,lang,1,objRIS.getEditable()))								
							%>
						</div>

						<div id="baths" style="float:left;padding-left:20px;height:40px;">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.baths")%></span><br/>
							<select name="field_baths__<%=idbat%>" id="field_baths">
							<option value=""></option>
							</select>
						</div>
						
						<div class="clear" style="height:10px;"></div>

						<div id="price">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.price")%></span><br/>
							<div style="height:35px;float:left;font-weight:bold;padding-left:3px;padding-right:3px;">da</div>
							<div style="height:35px;float:left;"><input type="text" style="width:100px;" name="price_from" id="price_from" value="<%=prcsxVal%>" onkeypress="javascript:return isInteger(event);"></div>
							<div style="height:35px;float:left;font-weight:bold;padding-left:3px;padding-right:3px;">a</div>
							<div style="height:35px;"><input type="text" style="width:100px;" name="price_to" id="price_to" value="<%=prcdxVal%>" onkeypress="javascript:return isInteger(event);"></div>
							<input type="hidden" value="" name="field_price__<%=idprc%>" id="field_price">							
						</div>

						<div id="superficie">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.superficie")%></span><br/>
							<div style="height:35px;float:left;font-weight:bold;padding-left:3px;padding-right:3px;">da</div>
							<div style="height:35px;float:left;"><input type="text" style="width:100px;" name="superficie_from" id="superficie_from" value="<%=supsxVal%>" onkeypress="javascript:return isInteger(event);"></div>
							<div style="height:35px;float:left;font-weight:bold;padding-left:3px;padding-right:3px;">a</div>
							<div style="height:35px;"><input type="text" style="width:100px;" name="superficie_to" id="superficie_to" value="<%=supdxVal%>" onkeypress="javascript:return isInteger(event);"></div>
							<input type="hidden" value="" name="field_superficie__<%=idsup%>" id="field_superficie">							
						</div>

						<div id="locali">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.locali")%></span><br/>
							<div style="height:35px;float:left;font-weight:bold;padding-left:3px;padding-right:3px;">da</div>
							<div style="height:35px;float:left;"><input type="text" style="width:100px;" name="locali_from" id="locali_from" value="<%=locsxVal%>" onkeypress="javascript:return isInteger(event);"></div>
							<div style="height:35px;float:left;font-weight:bold;padding-left:3px;padding-right:3px;">a</div>
							<div style="height:35px;"><input type="text" style="width:100px;" name="locali_to" id="locali_to" value="<%=locdxVal%>" onkeypress="javascript:return isInteger(event);"></div>
							<input type="hidden" value="" name="field_locali__<%=idloc%>" id="field_locali">							
						</div>

						<div id="accessori">
							<!--<div style="height:35px;float:left;">
							<input type="checkbox" name="field_accessori__<%=idacc%>" id="field_arredato" value="arredato">
							<span><%=lang.getTranslated("frontend.template.annunci.label.arredato")%></span>
							</div>
							<div style="height:35px;">
							<input type="checkbox" name="field_accessori__<%=idacc%>" id="field_giardino" value="giardino">
							<span><%=lang.getTranslated("frontend.template.annunci.label.giardino")%></span>
							</div>		

							<div style="height:35px;float:left;">
							<input type="checkbox" name="field_accessori__<%=idacc%>" id="field_terrazzo" value="terrazzo">
							<span><%=lang.getTranslated("frontend.template.annunci.label.terrazzo")%></span>
							</div>
							<div style="height:35px;">
							<input type="checkbox" name="field_accessori__<%=idacc%>" id="field_balcone" value="balcone">
							<span><%=lang.getTranslated("frontend.template.annunci.label.balcone")%></span>
							</div>-->	
							<%							
							fieldValueMatch = objListPairKeyValue(CStr(idacc))
							response.write(objContentF.renderContentFieldHTML(objACC,"", "__", "", fieldValueMatch,lang,1,objACC.getEditable()))								
							%>							
						</div>

						<div id="regione_provincia" style="padding-top:30px;">
							<span style="font-weight:bold;"><%=lang.getTranslated("frontend.template.annunci.label.regione_provincia")%></span><br/>
							<%							
							fieldValueMatch = objListPairKeyValue(CStr(idrpv))
							response.write(objContentF.renderContentFieldHTML(objRPV,"", "__", "", fieldValueMatch,lang,1,objRPV.getEditable()))								
							%>
						<script>
						
						$('#contentfield__<%=idrpv%> option').each(function() {
							var tmpv = $(this).val();
							if(tmpv.length>0 && tmpv.indexOf('IT-')==-1){
								$(this).remove();
							}
						});
						
						$('#contentfield__<%=idrpv%>').change(function() {
							$('#maplist').show();
							var tmpval = $('#contentfield__<%=idrpv%>').val();
							baseForZoom = tmpval.substring(0,tmpval.indexOf("_"));
							zoomLevel = 8;
							if(baseForZoom.length>5){
								zoomLevel = 10;
							}
							tmpval = tmpval.substring(tmpval.indexOf("_")+1);
							ajaxLoadLatLng(map, tmpval,zoomLevel);
								
							if($('#contentfield__<%=idrpv%>').val()!= "" && !hasGeoSearchActive){
								enableDrawingMode();	
							}else{
								disableDrawingMode();
							}
						});
						
						</script>
						</div>

						
					</form>			
					
					<div class="clear" style="height:10px;"></div>
					<div id="geosearchbuttons" style="float:left;"><input type="button" value="<%=lang.getTranslated("frontend.template.annunci.label.search")%>" onclick="javascript:sendGeoSearch();"/></div>
					<div id="searchresetbuttons" style="float:left;"><input type="button" value="<%=lang.getTranslated("frontend.template.annunci.label.reset_search")%>" onclick="javascript:reActivateDrawingMode();"/></div>
					<div id="georesetbuttons"><input type="button" value="<%=lang.getTranslated("frontend.template.annunci.label.clear_map")%>" onclick="javascript:resetGeoSearch();"/></div>					
					</div>
					<div class="clear" style="height:10px;"></div>
					<div style="width:480px;height:400px;vertical-align:top;text-align:left;display:none;border:1px solid;background:#FFFFFF;margin-bottom:30px;margin-top:30px;" id="maplist"></div>
			
					<script>
					jQuery(document).ready(function(){
						//ajaxLoadFilter("field_contract", "contratto", <%=langIdTarget%>, "", "<%=strParamPagFilter%>");				
						//ajaxLoadFilter("field_category", "categoria", <%=langIdTarget%>, "", "<%=strParamPagFilter%>");					
						//ajaxLoadFilter("field_typology", "tipologia", <%=langIdTarget%>, "", "<%=strParamPagFilter%>");
						//ajaxLoadFilter("field_typeproperty", "tipo-proprieta", <%=langIdTarget%>, "", "<%=strParamPagFilter%>");
						//ajaxLoadFilter("field_status", "stato", <%=langIdTarget%>, "", "<%=strParamPagFilter%>");
						//ajaxLoadFilter("field_riscaldamento", "riscaldamento", <%=langIdTarget%>, "", "<%=strParamPagFilter%>");
						ajaxLoadFilter("field_baths", "bagni", <%=langIdTarget%>, "", "<%=strParamPagFilter%>");
					});							
					</script>

				<%
				'************** codice per la lista news e paginazione
				'response.write("objListPoint.count: "&objListPoint.count&"<br>")
				'response.write("bolHasObj: "& bolHasObj &"<br>")
				if(bolHasObj) then%>
					<div id="content_list_container">
					<%for newsCounter = FromNews to ToNews
						Set objSelNews = objTmpNews(newsCounter)
						detailURL = "#"
						if(bolHasDetailLink) then
							detailURL = objMenuFruizione.resolveHrefUrl(base_url, (modelPageNum+1), lang, objCategoriaTmp, objTemplateSelected, objPageTempl)
						end if%>
						
						<div id="prodotto-immagine">
						<%if not(isNull(objSelNews.getFilePerNews())) AND not(isEmpty(objSelNews.getFilePerNews())) then
							Dim hasNotSmallImg
							hasNotSmallImg = true
							Set objListaFilePerNews = objSelNews.getFilePerNews()			
							for each xObjFile in objListaFilePerNews
								Set objFileXNews = objListaFilePerNews(xObjFile)
								iTypeFile = objFileXNews.getFileTypeLabel()
								if(Cint(iTypeFile) = 1) then%>	
									<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" alt="<%=objSelNews.getTitolo()%>" width="140" height="130" />
									<%hasNotSmallImg = false
									Exit for
								end if
								Set objFileXNews = nothing	
							next		
							if(hasNotSmallImg) then%>
								<img width="140" height="130" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0">
							<%end if
							Set objListaFilePerNews = nothing
						else%>
							<img width="140" height="130" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0">
						<%end if%>          
						</div>						
						
						<div id="prodotto-testo"><p class="title_contenuti"><a href="javascript:openDetailContentPage('<%=detailURL%>', '<%=strGerarchia%>', <%=objSelNews.getNewsID()%>, <%=(modelPageNum+1)%>);"><%=objSelNews.getTitolo()%></a></p>
						<%if (Len(objSelNews.getAbstract1()) > 0) then response.Write(objSelNews.getAbstract1()) end if%>
						</div>
						<div id="clear"></div>
						<div id="prodotto-footer"></div>
						<%Set objSelNews = nothing
					next%>
					<div><%if(totPages > 1) then call PaginazioneFrontend(totPages, numPage, strGerarchia, request.ServerVariables("URL"), strParamPagFilter) end if%></div>
					</div>
				<%end if%>
				</div>
			</div>
			<form action="" method="post" name="form_detail_link_news">	
			<input type="hidden" value="" name="id_news">	
			<input type="hidden" value="" name="modelPageNum">	
			<input type="hidden" value="" name="gerarchia">	
			<input type="hidden" value="<%=numPage%>" name="page">
			<input type="hidden" value="<%=order_by%>" name="order_by">            
			</form>
		</div>
		<!-- #include virtual="/public/layout/include/menu_vert_dx.inc" -->
	</div>
	<!-- #include virtual="/public/layout/include/bottom.inc" -->
</div>
</body>
</html>
<%
'****************************** PULIZIA DEGLI OGGETTI UTILIZZATI
Set objCategory = nothing
Set objPageTempl = nothing
Set objTemplate = nothing
Set objMenuFruizione = nothing
Set objListPoint = nothing
Set objListaTargetCat = nothing
Set objListaTargetLang = nothing
Set objListaNews = nothing
Set News = Nothing
%>
