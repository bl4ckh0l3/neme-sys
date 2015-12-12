<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%

Dim idComment, objCommento

'*** verifico se è stata passata la lingua dell'utente e la imposto come lang.setLangCode(xxx)
if not(isNull(request("lang_code"))) AND not(request("lang_code")="") AND not(request("lang_code")="null")  then
	lang.setLangCode(request("lang_code"))
	lang.setLangElements(lang.getListaElementsByLang(lang.getLangCode()))
end if

idComment = request("idComment")
Set objCommento = New CommentsClass
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<link rel="stylesheet" href="<%="http://" & request.ServerVariables("SERVER_NAME") & Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
</head>
<body>
	<div>
	<h2><%=lang.getTranslated("frontend.confirm_comment.mail.label.intro")%></h2>
	<br><br>


	<%
	Dim hasComments
	hasComments = false
	on error Resume Next

	Set objTmpCommento = objCommento.findCommentiByIDCommento(idComment,null,null)	
	
	if(Instr(1, typename(objTmpCommento), "CommentsClass", 1) > 0) then
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
		
		id_element = objTmpCommento.getIDElement()
		if(element_type=1) then
			Set objTmpElement = objNews.findNewsByID(id_element)%>
			<%=objTmpElement.getTitolo()%><br/><br/>
		<%else
			Set objTmpElement = objProd.findProdottoByID(id_element,false)%>
			<%=objTmpElement.getNomeProdotto()%><br/><br/>			
		<%end if
		Set objTmpElement = nothing%>
		<%=objTmpCommento.getDtaInserimento()%><br/>
		<%="<i>"&objUser.findUserByID(objTmpCommento.getIDUtente()).getUserName()&"</i>"%>
		<br><%=objTmpCommento.getMessage()%><br><hr>
		<%
		Set objUser = nothing
		Set objProd = nothing
		Set objNews = nothing%>

		<br/><br/><a href="<%="http://" & request.ServerVariables("SERVER_NAME") &Application("baseroot") &"/common/include/confirmcomment.asp?id_commento="&idComment%>"><%=lang.getTranslated("backend.confirm_comment.mail.label.confirm")%></a><br/><br/>		
	<%end if%>

	</div>
</body>
</html>
<%Set objCommento = nothing%>