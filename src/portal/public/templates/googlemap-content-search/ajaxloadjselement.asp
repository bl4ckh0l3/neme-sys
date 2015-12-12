<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
id = request("id")
result=""
Set objLoc = new LocalizationClass 
On Error Resume Next
Set objList = objLoc.findPointByElement(id,3)
if(strComp(typename(objList), "Dictionary", 1) = 0) then
	result = "["
	for each x in objList
	result=result&"new google.maps.LatLng("&objLoc.revertDoubleDelimiter(objList(x).getLatitude())&","&objLoc.revertDoubleDelimiter(objList(x).getLongitude())&"),"
	next
	result=result&("]")
	result=Replace(result, ",]", "]", 1, -1, 1)
end if
if(Err.number<>0)then
	response.write(Err.description)
	result=""
end if
response.write(result)
Set objLoc = nothing
%>