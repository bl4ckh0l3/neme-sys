<%
if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp, objListaTargetPerUser,numMaxImg
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("guest_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if

if not(isNull(objUserLogged.getTargetPerUser(objUserLogged.getUserID()))) then
	Set objListaTargetPerUser = objUserLogged.getTargetPerUser(objUserLogged.getUserID())	
end if
Set objUserLoggedTmp = nothing
Set objUserLogged = nothing

numMaxImg = Application("num_max_attachments")
if(not(request("numMaxImgs") = "")) then
	numMaxImg = request("numMaxImgs")
end if

'/**
'* recupero i valori della news selezionata se id_news <> -1
'*/
Dim id_news, strTitolo, strAbs1, strAbs2, strAbs3, strText, strKeyword, dtData_ins, dtData_pub, dtData_del, stato_news, objTarget, objFiles
Dim page_title, meta_description, meta_keyword
id_news = request("id_news")
strTitolo = ""
strAbs1 = ""
strAbs2 = ""
strAbs3 = ""
strText = ""
strKeyword = ""
dtData_ins = ""
dtData_pub = ""
dtData_del = ""
stato_news = -1
page_title = ""
meta_description = ""
meta_keyword = ""
objTarget = null
objFiles = null
%>