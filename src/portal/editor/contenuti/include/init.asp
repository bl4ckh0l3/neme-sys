<%
if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objCMSUtenteLogged"))
Set objUserLoggedTmp = nothing
Dim strRuoloLogged
strRuoloLogged = objUserLogged.getRuolo()
if not(strComp(Cint(strRuoloLogged), Application("admin_role"), 1) = 0) AND not(strComp(Cint(strRuoloLogged), Application("editor_role"), 1) = 0) then
	response.Redirect(Application("baseroot")&Application("error_page")&"?error=002")
end if
Set objUserLogged = nothing

Dim objNews, objListaNews, objContentField, objListaField
Set objNews = New NewsClass
Set objContentField = new ContentFieldClass

Dim order_news_by

if not(request("order_by") = "") then
	session("order_news_by") = request("order_by")
	order_news_by = session("order_news_by")
else
	if not(session("order_news_by") = "") then
		order_news_by = session("order_news_by")
	else
		session("order_news_by") = 1
		order_news_by = session("order_news_by")
	end if
end if

Dim totPages, itemsXpageNews, numPageNews,itemsXpageField, numPageField

showTab="contenutilist"
if(request("showtab")<>"")then
	showTab=request("showtab")
end if

if not(request("itemsNews") = "") then
	session("contenutiItems") = request("itemsNews")
	itemsXpageNews = session("contenutiItems")
	session("contenutiPage") = 1
else
	if not(session("contenutiItems") = "") then
		itemsXpageNews = session("contenutiItems")
	else
		session("contenutiItems") = 20
		itemsXpageNews = session("contenutiItems")
	end if
end if

if (showTab="contenutilist") AND not(request("page") = "") then
	session("contenutiPage") = request("page")
	numPageNews = session("contenutiPage")
else
	if not(session("contenutiPage") = "") then
		numPageNews = session("contenutiPage")
	else
		session("contenutiPage") = 1
		numPageNews = session("contenutiPage")
	end if
end if	


if not(request("itemsField") = "") then
	session("fieldItems") = request("itemsField")
	itemsXpageField = session("fieldItems")
	session("fieldPage") = 1
else
	if not(session("fieldItems") = "") then
		itemsXpageField = session("fieldItems")
	else
		session("fieldItems") = 20
		itemsXpageField = session("fieldItems")
	end if
end if

if (showTab="contenutifield") AND not(request("page") = "") then
	session("fieldPage") = request("page")
	numPageField = session("fieldPage")
else
	if not(session("fieldPage") = "") then
		numPageField = session("fieldPage")
	else
		session("fieldPage") = 1
		numPageField = session("fieldPage")
	end if
end if


Dim target_cat_param
target_cat_param = ""

Dim CategoriatmpClass, objListCatXNews
Set CategoriatmpClass = new CategoryClass
%>