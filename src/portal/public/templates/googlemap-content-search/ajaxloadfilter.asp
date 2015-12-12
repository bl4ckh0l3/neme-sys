<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->

<%
description = request("description")
target_lang = request("target_lang")
sorting = request("sorting")

Set objContentF = new ContentFieldClass
Set ajaxObjListPairKeyValue = Server.CreateObject("Scripting.Dictionary")	

if (Request.Form.Count>0)then      
	For Each y In Request.Form
		if (Instr(1, y, "field_", 1) > 0) then
			if(Trim(Request.Form(y))<>"") then
				  tmpKey = Mid(y,Instr(1, y, "__", 1)+2)
				  tmpValue = Request.Form(y)
				  ajaxObjListPairKeyValue.add tmpKey,tmpValue
			end if
		end if
	Next 
end if   

On Error Resume Next
Set objList = objContentF.getListContentFieldValuesByDesc(description, target_lang, sorting)
for each k in objList
	label=k
	if not(lang.getTranslated(label)="") then label=lang.getTranslated(label) end if%>
	<option value="<%=k%>" <%if(strComp(ajaxObjListPairKeyValue(Cstr(objList(k))), k, 1) = 0)then response.write("selected") end if%>><%=label%></option>
<%next
if(Err.number<>0)then
'response.write(Err.description)
end if

Set ajaxObjListPairKeyValue = nothing
Set objContentF = nothing%>