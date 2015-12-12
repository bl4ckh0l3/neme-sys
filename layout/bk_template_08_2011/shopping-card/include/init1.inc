<%
Dim id_user, scontoCliente, objCarrelloUser, carrello, Prodotto, hasSessionIDCard, hasSconto, objTasse, hasGroup, groupCliente, groupDesc
Dim objGroup
Set objGroup = new UserGroupClass

hasSconto=false
hasGroup = false
scontoCliente = 0
groupCliente = ""
groupDesc = ""

Set carrello = New CardClass
Set Prodotto = New ProductsClass
Set objTasse = new TaxsClass

Set objProdField = new ProductFieldClass
Set objProdFieldGroup = new ProductFieldGroupClass
Set objShip = new ShippingAddressClass

objCarrelloUser = null

if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objUserLoggedTmp = nothing
	
	if(objUserLogged.getRuolo() <> 3) then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error=023")
	end if	
		
	hasSessionIDCard = carrello.findCarrelloByIDUser(Session.SessionID)	
	id_user = objUserLogged.getUserID()

	  groupCliente = objUserLogged.getGroup()
	  if(not(groupCliente= "")) then
      On Error Resume Next
      Set objTmpGr = objGroup.findUserGroupByID(groupCliente)
      groupDesc = objTmpGr.getShortDesc()
	  if (not(isNull(objTmpGr.getTaxGroup()))) then
		  Set groupClienteTax = objTmpGr.getTaxGroupObj(objTmpGr.getTaxGroup())
	  end if
      hasGroup = true
      Set objTmpGr = nothing
      Set objSelMargin = objGroup.getMarginDiscountXUserGroup(groupCliente)
      if(Err.number <> 0) then
        hasGroup = false
      end if
	  end if

	scontoCliente= objUserLogged.getSconto()
	
	if(scontoCliente <> "" AND Cdbl(scontoCliente) > 0) then 
		hasSconto = true
	end if
			
	if(hasSessionIDCard) then	
		Set objCarrelloUser = carrello.getCarrelloByIDUser(Session.SessionID)	
		call objCarrelloUser.updateIDUtenteCarrello(objCarrelloUser.getIDCarrello(), id_user)
		Set objCarrelloUser = carrello.getCarrelloByIDUser(id_user)
	else
		Set objCarrelloUser = carrello.getCarrelloByIDUser(id_user)
	end if
	
	Set objUserLogged = nothing
end if	

Dim objListaCarrello, order_carrello_by
Dim id_prodotto, operation, qta_prod, prezzo, totale_prod, id_carrello
order_carrello_by = 3

if(carrello.findCarrelloByIDUser(Session.SessionID) = true) then
	Set objCarrelloUser = carrello.getCarrelloByIDUser(Session.SessionID)
end if

'**************** IMPOSTO LE CLASSI PER LA VALUTA
Dim currClass, defCurrObj, hasCurrency
Set currClass = new CurrencyClass
On Error Resume Next
hasCurrency = false
defCurrObj = currClass.getDefaultCurrency().getCurrency()
if(Err.number <> 0) then
	defCurrObj = ""
end if

if not(defCurrObj="") AND not(Session("currency")="") then
	hasCurrency = true
end if


'********** GESTIONE INTERNAZIONALIZZAZIONE TASSE
Dim international_country_code, international_state_region_code
international_country_code = ""
international_state_region_code = ""

On Error Resume Next
Set orderShip = objShip.findShippingAddressByUserID(id_user)

if (Instr(1, typename(orderShip), "ShippingAddressClass", 1) > 0) then
	international_country_code = orderShip.getCountry()	
	if not(isNull(orderShip.getStateRegion()) AND not(orderShip.getStateRegion()="")) then
		international_state_region_code = orderShip.getStateRegion()
	end if	
	hasShipAddress = true
end if
Set orderShip = nothing

if(Err.number <> 0) then 
'response.write(Err.description)
end if
%>