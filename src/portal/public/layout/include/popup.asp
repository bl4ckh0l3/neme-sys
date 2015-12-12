<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
'<!--nsys-popup1-->
%>
<!-- #include virtual="/common/include/Objects/File4ProductsClass.asp" -->
<%
'<!---nsys-popup1-->
%>
<%
Dim id_allegato, name_allegato, type_allegato, path_allegato, dida_allegato, objFiles, objSelectedFile, parent_type
id_allegato = request("id_allegato")
parent_type = request("parent_type")

if(parent_type=1) then
	Set objFiles = new File4NewsClass
'<!--nsys-popup2-->
elseif(parent_type=2) then
	Set objFiles = new File4ProductsClass
'<!---nsys-popup2-->
end if

Set objSelectedFile = objFiles.getFileByID(id_allegato)
Set objFiles = nothing

name_allegato = objSelectedFile.getFileName()
type_allegato = objSelectedFile.getFileTypeLabel()

if(parent_type="1") then
	path_allegato = Application("dir_upload_news")&objSelectedFile.getFilePath()
'<!--nsys-popup3-->
elseif(parent_type="2") then
	path_allegato = Application("dir_upload_prod")&objSelectedFile.getFilePath()
'<!---nsys-popup3-->
end if
dida_allegato = objSelectedFile.getFileDida()

' LEGENDA TIPI FILE
'1 = img small
'2 = img big
'3 = pdf
'4 = audio-video
'5 = others
'6 = img medium
'7 = img card
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
</head>
<body>
	<div id="container">	
		<div id="content-popup">
		<%
		select case Cint(type_allegato)
		case 4%>				
			
		<object id="MediaPlayer" classid="CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95" width="300" height="300" codebase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715" type="application/x-oleobject">
		  <param name="FileName" value="<%=path_allegato%>">
			<param name="DisplaySize" value="0">
			<param name="AnimationatStart" value="true">
			<param name="TransparentatStart" value="true">
			<param name="AutoStart" value="true">
			<param name="loop" value="false">
			<param name="AllowChangeDisplaySize" value="true">
			<param name="AutoSize" value="false">
			<param name="ShowControls" value="true">
			<param name="ShowStatusBar" value="true">
			<param name="EnablePositionControls" value="false">
			<param name="ShowTracker" value="false">
			<param name="ShowPositionControls" value="false">
		  <embed type="application/x-mplayer2" pluginspage="http://www.microsoft.com/isapi/redir.dll?prd=windows&sbp=mediaplayer&ar=Media&sba=Plugin&" name=\"MediaPlayer\" filename="<%=path_allegato%>" width="300" height="300" loop="0" animationatstart="1" showstatusbar="1" transparentatstart="1" showcontrols="1" displaysize="0" enablepositioncontrols="0" showTracker="0" showpositioncontrols="0" />
		</object>	
		<%case 1, 2, 6, 7%>
			<img src="<%=path_allegato%>" border="0" vspace="2" hspace="2" alt="<%=dida_allegato%>">
		<%case else%>
			<%=lang.getTranslated("frontend.popup.label.download_selected_file")%>: <a target="_blank" href="<%=path_allegato%>"><%=name_allegato%></a>
		<%end select%>
		<div align="center" style="padding-top:30px;">	
		<a href="javascript:window.close();"><%=lang.getTranslated("frontend.popup.label.close_window")%></a></div>
		</div>
	</div>
</body>
</html>