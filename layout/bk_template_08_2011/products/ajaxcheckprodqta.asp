<!-- #include virtual="/editor/include/IncludeShopObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/CardClass.asp" -->
<!-- #include virtual="/common/include/Objects/ProductsCardClass.asp" -->
<!-- #include virtual="/common/include/Objects/ProductFieldClass.asp" -->
<%

Dim carrello, objCardProds, objProdPerCarrello, objProdField
Dim id_card, id_prod
Set carrello = New CardClass
Set objProdPerCarrello = New ProductsCardClass
Set objProdField = New ProductFieldClass
id_card = -1
id_prod = request("id_prod")
prod_counter = request("prod_counter") 
qta_checked = 0

if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objUserLoggedTmp = nothing
	
	if(objUserLogged.getRuolo() <> 3) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=023")
	end if
	
	id_user = objUserLogged.getUserID() 
	
	'verifico esistenza carrello e recupero lista prodotti + quantità per vedere se l'utente vuole comprare qta diverse da 1 o multipli 6
	hasCard = carrello.findCarrelloByIDUser(id_user)
	if(hasCard)then
		id_card = carrello.getCarrelloByIDUser(id_user).getIDCarrello()
		Set objCardProdList = objProdPerCarrello.getListItem(id_card,id_prod)
	else		
		hasCard = carrello.findCarrelloByIDUser(Session.SessionID)
		if(hasCard)then
			id_card = carrello.getCarrelloByIDUser(Session.SessionID).getIDCarrello()
			Set objCardProdList = objProdPerCarrello.getListItem(id_card,id_prod)		
		end if
	end if
	
	Set objUserLogged = nothing
else	
	hasCard = carrello.findCarrelloByIDUser(Session.SessionID)
	if(hasCard)then
		id_card = carrello.getCarrelloByIDUser(Session.SessionID).getIDCarrello()
		Set objCardProdList = objProdPerCarrello.getListItem(id_card,id_prod)	
	end if
end if


On Error Resume Next
selectedCounter = 0
hasFieldProdCardCombination = ""	

Set fieldList4Card = objProdField.findListFieldXCardByProd(null, id_card, id_prod)
if(fieldList4Card.count > 0)then											
	for each k in fieldList4Card	
		hasFieldProdCardCombination = ""							
		Set objTmpField4Card = fieldList4Card(k)
		keys = objTmpField4Card.Keys
		
		for each r in keys
			Set tmpF4O = r
			field_prod_val = request(objProdField.getFieldPrefix()&prod_counter&tmpF4O.getID())
			if(Trim(field_prod_val)<>"")then
				if(Trim(tmpF4O.getSelValue())=Trim(field_prod_val))then
					if(hasFieldProdCardCombination = "")then
						hasFieldProdCardCombination = true
					else
						if(hasFieldProdCardCombination)then
							hasFieldProdCardCombination = true
						else
							hasFieldProdCardCombination = false
						end if										
					end if	
				end if
			end if
			Set tmpF4O =nothing
		next	

		if (hasFieldProdCardCombination <> "") AND (hasFieldProdCardCombination=true)then	
			qta_checked = objCardProdList.Item(id_prod&"|"&k).getQtaProd()
			Exit for
		end if
		
		Set objTmpField4Card = nothing
	next
end if								
Set fieldList4Card = nothing
if(Err.number<>0)then
	qta_checked = 0
end if

response.write(qta_checked)

Set objProdField = nothing
Set objProdPerCarrello = nothing
Set carrello = nothing
%>
