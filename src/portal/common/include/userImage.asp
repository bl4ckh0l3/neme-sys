<%@LANGUAGE="VBSCRIPT"%>
<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
Dim thisFileName, thisContentType, thisfileData, objUser, userID, objUserImage, userImageData

userID = request("userID")

Set objUser = new UserClass
Set objUserImage = objUser.getUserImageObjectNoData(userID)
userImageData = objUser.getUserImageData(userID)
thisFileName = objUserImage.item("filename")
thisContentType = objUserImage.item("content_type")
thisfileData = userImageData
Set objUserImage = nothing
Set objUser = nothing

Sub sendUserImage()
	Response.Clear
	Response.Buffer = True
	Response.ContentType = thisContentType
	Response.AddHeader "Content-Disposition", "inline; filename="&thisFileName
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	Response.BinaryWrite(thisfileData)		
	Response.Flush			
End Sub
%>
<%sendUserImage()%>