<%@LANGUAGE="VBSCRIPT"%>
<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/UserFilesClass.asp" -->

<%
' Impostazione che setta il tipo di file in output su XML
Response.Buffer = TRUE
response.expires = -1500 
response.ContentType = "text/xml"
response.flush

Dim objFiles, userID, showDel

userID = request("userID")

showDel = "0"
if (Cint(userID) = Session("objUtenteLogged")) then
	showDel = "1"
end if


Set objFiles = new UserFilesClass

Dim hasImg, objListaPhotos
hasImg = false

Set objListaPhotos = objFiles.getFiles4User(userID)	

if(objListaPhotos.Count > 0) then
hasImg = true
end if

if Err.number <> 0 then
'response.write(Err.description)
end if	
%>
<tiltviewergallery>
	<photos>

	<%if(hasImg) then	
		iIndex = objListaPhotos.Count	

		objTmpImg = objListaPhotos.Items				

		for photoCounter = 0 to iIndex-1
		Set objFilteredPhoto = objTmpImg(photoCounter)
		%>
		<photo imageurl="<%=Application("baseroot") & Application("dir_upload_user")&objFilteredPhoto.getFilePath()%>"<%if(showDel = "1")then%> linkurl="<%=Application("baseroot")&"/area_user/delphoto.asp?id_user="&userID&"&amp;closewindow=1&amp;id_photo="&objFilteredPhoto.getFileID()%>"<%end if%>>
			<title><%=objFilteredPhoto.getFileName()%></title>
			<description><![CDATA[<%=objFilteredPhoto.getFileDida()%><br/><br/><%=FormatDateTime(objFilteredPhoto.getDataIns(),2)%>]]></description>
		</photo>
		
		<!--<div>
		<img id="user_photo_<%'=photoCounter%>" src="<%'=Application("baseroot") & Application("dir_upload_user")&objFilteredPhoto.getFilePath()%>" align="left" width="200"/><br/>
		<script>resizeimagesByID('user_photo_<%'=photoCounter%>', 200);</script>
		<%'=objFilteredPhoto.getFileDida()%><br/>
		<%'=FormatDateTime(objFilteredPhoto.getDataIns(),2)%><br/>
		<a title="<%'=lang.getTranslated("portal.templates.commons.label.del_friend")%>" href="javascript:delPhoto(<%'=objFilteredPhoto.getFileID()%>,<%'=objFilteredPhoto.getUserID()%>);"><img id="del" src="<%'=Application("baseroot") & "/common/img/cancel.png"%>"/></a>
		</div>-->
		<%
		Set objFilteredPhoto = nothing
		next
		Set objTmpImg = nothing
		Set objListaPhotos = nothing
		Set objFiles = nothing
	end if%>   

	</photos>
</tiltviewergallery>

<%
Set objFiles = nothing
%>