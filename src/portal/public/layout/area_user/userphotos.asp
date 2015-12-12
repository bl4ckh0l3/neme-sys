<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/area_user.css"%>" type="text/css">
<script type="text/JavaScript" src="<%=Application("baseroot") & "/common/js/swfobject.js"%>"></script>
<script language="JavaScript">

function delPhoto(idphoto,idUser){
	if(confirm("<%=lang.getTranslated("frontend.area_user.manage.label.delphoto")%>")){
		location.href='<%=Application("baseroot") & "/area_user/delphoto.asp?id_photo="%>'+idphoto+'&id_user='+idUser;
	}
}

function changeTab(number){
	if(number==1)
		location.href='<%=Application("baseroot") & "/area_user/userprofile.asp"%>';
	else if(number==2)
		location.href='<%=Application("baseroot") & "/area_user/manageuser.asp"%>';
	else if(number==3)
		location.href='<%=Application("baseroot") & "/area_user/friendlist.asp"%>';
	else if(number==4)
		location.href='<%=Application("baseroot") & "/area_user/userphotos.asp"%>';

}
  
function changeNumMaxImgs(){
	if(document.form_inserisci.numMaxImgs.value == ""){
		alert("<%=lang.getTranslated("frontend.area_user.js.alert.insert_value")%>");
		document.form_inserisci.numMaxImgs.focus();
		return;
	}else if(isNaN(document.form_inserisci.numMaxImgs.value)){
		alert("<%=lang.getTranslated("frontend.area_user.js.alert.isnan_value")%>");
		document.form_inserisci.numMaxImgs.focus();
		return;		
	}
	location.href = "<%=Application("baseroot") & "/area_user/userphotos.asp?numMaxImgs="%>"+document.form_inserisci.numMaxImgs.value;
}
  
  function sendForm(){
	//if(controllaCampiInput()){

		<%if(Application("use_aspupload_lib") = 1) then%>
			document.form_inserisci.action="<%=Application("baseroot") & "/area_user/ProcessPhoto2.asp"%>";
		<%end if%>
		
		document.getElementById("loading").style.visibility = "visible";
		document.getElementById("loading").style.display = "block";
		document.form_inserisci.submit();
	//}else{
	//	return;
	//}
}
</script>
<!-- #include virtual="/common/include/initCommonJs.inc" -->
</head>
<body>
<!-- #include file="grid_top.asp" -->

        <h1><%=lang.getTranslated("frontend.header.label.utente_photo")%>&nbsp;<em><%=strUserName%></em></h1>
        <p>
		<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.profile")%>" type="button" onclick="javascript:changeTab(1);">
		<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.modify")%>" type="button" onclick="javascript:changeTab(2);">
		<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.friends")%>" type="button" onclick="javascript:changeTab(3);">
		<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.photos")%>" type="button" onclick="javascript:changeTab(4);">
		</p>
		<div id="profilo-utente">        


		<form action="<%=Application("baseroot") & "/area_user/ProcessPhoto.asp"%>" method="post" name="form_inserisci" enctype="multipart/form-data">
		<input type="hidden" value="<%=id_utente%>" name="id_utente">

		<%
		Dim objFiles, objListFileType
		Set objFiles = new UserFilesClass
		Set objListFileLabel = objFiles.getListaFilesLabel()
		%>
		  <%
		  Dim fileCounter
		  for fileCounter=1 to numMaxImg%>
		  
      <div align="left"><br/>
      <div style="float:left;"><%=lang.getTranslated("backend.contenuti.detail.table.label.attachment")%><br/>
			<input type="file" name="fileupload<%=fileCounter%>" class="formFieldTXT">
			</div>
      <div>
       <%if(fileCounter=1)then%><%=lang.getTranslated("backend.commons.detail.table.label.change_num_imgs")%>
      <br/> 
      <input type="text" value="<%=numMaxImg%>" name="numMaxImgs" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);"><a href="javascript:changeNumMaxImgs();"><img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="0" hspace="4" border="0" align="top" alt="<%=lang.getTranslated("backend.commons.detail.table.label.change_num_imgs")%>"></a>
      <%else%>
      <br/> <br/>         
      <%end if%>
      </div>
      <br/> 
      <div style="float:left;"><%=lang.getTranslated("backend.contenuti.detail.table.label.file_type_label")%><br/>
			<select name="fileupload<%=fileCounter%>_label" class="formFieldSelectTypeFile">
			<%for each xType in objListFileLabel%>
			<option value="<%=xType%>"><%=objListFileLabel(xType)%></option>
			<%next%>
			</select>
			</div>
      
			<div><%=lang.getTranslated("backend.contenuti.detail.table.label.file_dida")%><br/>
      <input type="text" name="fileupload<%=fileCounter%>_dida" class="formFieldTXT"> 
      </div>
		 <%next%> 	
      <br/> 
			<div>
      <input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=lang.getTranslated("backend.contenuti.detail.button.inserisci.label")%>" onclick="javascript:sendForm();" />
	  	</div>
      </div>

    </form>
    <br/><br/>

        <div id="flashcontent">TiltViewer requires JavaScript and the latest Flash player. <a href="http://www.macromedia.com/go/getflashplayer/">Get Flash here.</a></div>
        <script type="text/javascript">
        
          var fo = new SWFObject("<%=Application("baseroot")&"/common/swf/TiltViewer.swf"%>", "viewer", "100%", "400px", "9.0.28", "#000000");			
          
          // TILTVIEWER CONFIGURATION OPTIONS
          // To use an option, uncomment it by removing the "//" at the start of the line
          // For a description of config options, go to: 
          // http://www.airtightinteractive.com/projects/tiltviewer/config_options.html
                                    
          //FLICKR GALLERY OPTIONS
          // To use images from Flickr, uncomment this block
          //fo.addVariable("useFlickr", "true");
          //fo.addVariable("user_id", "48508968@N00");
          //fo.addVariable("tags", "jump,smile");
          //fo.addVariable("tag_mode", "all");
          //fo.addVariable("showTakenByText", "true");			
          
          // XML GALLERY OPTIONS
          // To use local images defined in an XML document, use this block		
          fo.addVariable("useFlickr", "false");
          fo.addVariable("xmlURL", "<%=Application("baseroot")&"/area_user/userphotoxml.asp?userID="&id_utente%>");
          fo.addVariable("maxJPGSize","640");
          
          //GENERAL OPTIONS		
          fo.addVariable("useReloadButton", "false");
          fo.addVariable("columns", "3");
          fo.addVariable("rows", "3");
          //fo.addVariable("showFlipButton", "true");
          //fo.addVariable("showLinkButton", "true");		
          fo.addVariable("linkLabel", "Delete image");
          //fo.addVariable("frameColor", "0xFF0000");
          //fo.addVariable("backColor", "0xDDDDDD");
          //fo.addVariable("bkgndInnerColor", "0xFF00FF");
          //fo.addVariable("bkgndOuterColor", "0x0000FF");				
          //fo.addVariable("langGoFull", "Go Fullscreen");
          //fo.addVariable("langExitFull", "Exit Fullscreen");
          //fo.addVariable("langAbout", "About");				
          
          // END TILTVIEWER CONFIGURATION OPTIONS
          
          fo.addParam("allowFullScreen","true");
          fo.write("flashcontent");			
        </script>	

      

      <div id="loading" style="visibility:hidden;display:none;" align="center"><img src="/editor/img/loading.gif" vspace="0" hspace="0" border="0" alt="Loading..." width="200" height="50"></div>


          <form name="refreshAfterDelPhoto" id="refreshAfterDelPhoto" method="post" action="<%=Application("baseroot") & "/area_user/userphotos.asp"%>">
            <input type="hidden" name="id_utente" id="id_utente" value="<%=id_utente%>">
          </form>
       </div>	
		   
<!-- #include file="grid_bottom.asp" -->
</body>
</html>