<!-- #include virtual="/common/include/Objects/LocalizationClass.asp" -->
<%
objtype = request("type")
vertices = request("vertices")
center = request("center")
radius = request("radius")
current_overlay = request("current_overlay")
last_selection= request("last_selection")

if (Instr(1, typename(Session("geolocalsearchpoly")), "Dictionary", 1) > 0) then
	Set objGeolocalSearch = Session("geolocalsearchpoly")
else
	Set objGeolocalSearch = Server.CreateObject("Scripting.Dictionary")
end if

objGeolocalSearch.item("type") = objtype
objGeolocalSearch.item("current_overlay") = current_overlay
objGeolocalSearch.item("last_selection") = last_selection
objGeolocalSearch.item("search_active") = "0"


if(Cint(objtype)=1)then
	objGeolocalSearch.item("vertices") = vertices	
	Set Session("geolocalsearchpoly") = objGeolocalSearch
elseif(Cint(objtype)=2)then
	objGeolocalSearch.item("center") = center
	objGeolocalSearch.item("radius") = radius
	Set Session("geolocalsearchpoly") = objGeolocalSearch
elseif(Cint(objtype)=3)then
	'objGeolocalSearch.remove("type")
	'objGeolocalSearch.remove("current_overlay")
	'objGeolocalSearch.remove("last_selection")
	'objGeolocalSearch.remove("vertices")
	'objGeolocalSearch.remove("center")
	'objGeolocalSearch.remove("radius")
	objGeolocalSearch.removeAll
	objGeolocalSearch.item("search_active") = "0"
	Set Session("geolocalsearchpoly") = objGeolocalSearch
else
	Session("geolocalsearchpoly") = null
end if

Set objGeolocalSearch = nothing
%>