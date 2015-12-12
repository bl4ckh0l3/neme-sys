<!-- #include file="IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/CommentsClass.asp" -->
<!-- #include virtual="/common/include/Objects/NewsClass.asp" -->
<!-- #include virtual="/common/include/Objects/File4NewsClass.asp" -->
<%
if (isEmpty(Session("objCMSUtenteLogged"))) then%>
	<html>
	<head>
	</head>
	<body onload="window.close();">
	</body>
	</html>
	<%response.end()
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
Set objUserLoggedTmp = nothing

Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then%>
	<html>
	<head>
	</head>
	<body onload="window.close();">
	</body>
	</html>
	<%response.end()
end if
Set objUserLogged = nothing


Dim element_type, objCommento, objListaComments, del_commento, id_commento, unlock_commento
element_type = request("element_type")
id_commento = request("id_commento")
del_commento = request("del_commento")
unlock_commento = request("unlock_commento")

Set objCommento = New CommentsClass

if(del_commento = 1) then
	call objCommento.deleteCommento(id_commento)
end if

if(unlock_commento = 1) then
	call objCommento.updateStatus(id_commento,1)
end if
%>	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="initCommonMeta.inc" -->
<!-- #include file="initCommonJs.inc" -->
<script language="JavaScript">
function sendForm(){
	document.send_commento.submit();	
}
</script>
</head>
<body>
	<div id="container">
		<div id="content-popup">	
		<%
		Dim hasComments
		hasComments = false
		on error Resume Next

			Set objListaComments = objCommento.findCommentiByType(element_type,null)	
			
			if(objListaComments.Count > 0) then
				hasComments = true
			end if
			
		if Err.number <> 0 then
			hasComments = false
		end if	
		
		if(hasComments) then
			Dim  x, objTmpCommento, id_element, objNews, objProd, objUser
			id_element = -1
			Set objNews = New NewsClass
			Set objProd = New ProductsClass
			Set objUser = New UserClass
			
			
			for each x in objListaComments.Keys
				Set objTmpCommento = objListaComments(x)
				if not(CInt(objTmpCommento.getIDElement()) = id_element) then
					id_element = objTmpCommento.getIDElement()
					if(element_type=1) then
						Set objTmpElement = objNews.findNewsByID(id_element)%>
						<a href="javascript:window.opener.editContent(<%=id_element%>);"><%=objTmpElement.getTitolo()%></a><br/><br/>
					<%else
						Set objTmpElement = objProd.findProdottoByID(id_element,false)%>
						<a href="javascript:window.opener.editProduct(<%=id_element%>);"><%=objTmpElement.getNomeProdotto()%></a><br/><br/>			
					<%end if
					Set objTmpElement = nothing
				end if%>
				<%=objTmpCommento.getDtaInserimento()%>
				<%if(objTmpCommento.getActive()=0) then%>
				<a href="<%=Application("baseroot") & "/editor/include/popupCommentManager.asp?unlock_commento=1&id_commento="&objTmpCommento.getIDCommento()&"&element_type="&element_type%>" title="<%=langEditor.getTranslated("portal.templates.commons.label.comment_unlock")%>"><img src="<%=Application("baseroot")&"/editor/img/lock_open.png"%>" vspace="0" hspace="2" border="0" align="absmiddle" alt="<%=langEditor.getTranslated("portal.templates.commons.label.comment_unlock")%>"></a>&nbsp;&nbsp;
				<%end if%>
				<a href="<%=Application("baseroot") & "/editor/include/popupCommentManager.asp?del_commento=1&id_commento="&objTmpCommento.getIDCommento()&"&element_type="&element_type%>"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" vspace="0" hspace="2" border="0" align="absmiddle"></a><br/>
				<%="<i>"&objUser.findUserByID(objTmpCommento.getIDUtente()).getUserName()&"</i>"%>
				<br><%=objTmpCommento.getMessage()%><br><hr>
			<%next	

			Set objUser = nothing
			Set objProd = nothing
			Set objNews = nothing
		else
			response.Write("<div align='center'>"&langEditor.getTranslated("frontend.popup.label.no_comments")&"</div><br>")
		end if%>
		<br><br>
		<div align="center" style="padding-top:30px;"><a href="javascript:window.close();"><%=langEditor.getTranslated("frontend.popup.label.close_window")%></a></div><br>
		</div>
	</div>
</body>
</html>
<%
Set objListaComments = nothing
Set objCommento = nothing%>