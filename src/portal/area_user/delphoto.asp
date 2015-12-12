<%On Error Resume Next%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/UserFilesClass.asp" -->
<%if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUsrFiles
Set objUsrFiles = new UserFilesClass%>
<%

Dim idUser,idPhoto
idUser = request("id_user")
idPhoto = request("id_photo")


Set objDB = New DBManagerClass
Set objConn = objDB.openConnection()
objConn.BeginTrans
call objUsrFiles.deleteFiles(idPhoto, idUser, objConn)

if objConn.Errors.Count = 0 then
	objConn.CommitTrans
else
	objConn.RollBackTrans
	response.Redirect(Application("baseroot") & "/area_user/userphotos.asp?del_done=0")
end if			
Set objDB = nothing	

if(request("closewindow")="1")then%>
<html>
<head>
    <script language="javascript" type="text/javascript">
      <!--

      // Funzione usata per aggiornare la pagina opener una volta cancellata la photo selezionata
      function refreshParentData() {
            opener.document.forms["refreshAfterDelPhoto"].submit();
      }

      //-->
    </script>      
</head>
<body onload="refreshParentData();window.close();">
</body>	
</html>	
<%else
	response.Redirect(Application("baseroot") & "/area_user/userphotos.asp?del_done=1")	
end if

Set objUsrFiles = nothing

if(Err.number <> 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if
%>