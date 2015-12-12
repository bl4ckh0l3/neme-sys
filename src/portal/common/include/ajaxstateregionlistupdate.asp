<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
	Dim field_name, field_val, objtype, id_objref
	field_val = request("field_val")

	Set objRef = New CountryClass
	if(field_val<>"")then
		On Error Resume Next
		Set specialFieldValue = objRef.findStateRegionListByCountry(field_val,"2,3")			    
		for each x in specialFieldValue
			key =  specialFieldValue(x).getStateRegionCode()%>
			<option value="<%=key%>"><%=Server.HTMLEncode(lang.getTranslated("portal.commons.select.option.country."&key))%></option>     
		<%next
		Set specialFieldValue = nothing
		if(Err.number <>0)then

		end if
	end if
	Set objRef = nothing
%>
