<%
Dim id_prodotto, Prodotto, objCurrentProdotto, strGerarchia, strCSS, objFileXProdotto, objListaFilePerProdotto, objSelMargin


Set Prodotto = New ProductsClass
id_prodotto = request("id_prodotto")
strGerarchia = request("gerarchia")

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

Set objCommento = New CommentsClass

'**************** IMPOSTO LE CLASSI PER LA VALUTA
Dim currClass, defCurrObj, thisCurrObj
Set currClass = new CurrencyClass
On Error Resume Next
defCurrObj = currClass.getDefaultCurrency().getCurrency()
if(Err.number <> 0) then
	defCurrObj = ""
end if

if(request("currency") <> "") then
	Session("currency") = request("currency")
	thisCurrObj = Session("currency")
elseif(Session("currency") <> "") then
	thisCurrObj = Session("currency")
else
	thisCurrObj = defCurrObj
	Session("currency") = thisCurrObj
end if

Set objProdField = new ProductFieldClass
Set objProdFieldGroup = new ProductFieldGroupClass

hasSconto=false
hasGroup = false
scontoCliente = 0
groupCliente = ""
groupDesc = ""
groupClienteTax = null	
objSelMargin = null	


'********** GESTIONE INTERNAZIONALIZZAZIONE TASSE
Dim international_country_code, international_state_region_code, userIsCompanyClient
international_country_code = ""
international_state_region_code = ""
userIsCompanyClient = 0
groupClienteTax = null

if not(isEmpty(Session("objUtenteLogged"))) then
	Dim objUserLogged, objUserLoggedTmp
	Set objUserLoggedTmp = new UserClass
	Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))
	Set objUserLoggedTmp = nothing

	id_user = objUserLogged.getUserID()

  groupCliente = objUserLogged.getGroup()
  if(not(groupCliente= "")) then
    On Error Resume Next
    Dim objGroup
    Set objGroup = New UserGroupClass
    Set objTmpGr = objGroup.findUserGroupByID(groupCliente)
    groupDesc = objTmpGr.getShortDesc()
    if (not(isNull(objTmpGr.getTaxGroup()))) then
      Set groupClienteTax = objTmpGr.getTaxGroupObj(objTmpGr.getTaxGroup())
    end if
    hasGroup = true
    Set objTmpGr = nothing
		Set objSelMargin = objGroup.getMarginDiscountXUserGroup(groupCliente)
    Set objGroup = nothing
    if(Err.number <> 0) then
      hasGroup = false
    end if
  end if

	scontoCliente= objUserLogged.getSconto()
	
	if(scontoCliente <> "" AND Cdbl(scontoCliente) > 0) then
		hasSconto = true
	end if
  
	Set objUserLogged = nothing


  On Error Resume Next
  Set objShip = new ShippingAddressClass
  Set orderShip = objShip.findShippingAddressByUserID(id_user)

  if (Instr(1, typename(orderShip), "ShippingAddressClass", 1) > 0) then
    international_country_code = orderShip.getCountry()	
    if not(isNull(orderShip.getStateRegion()) AND not(orderShip.getStateRegion()="")) then
      international_state_region_code = orderShip.getStateRegion()
    end if
    userIsCompanyClient = orderShip.isCompanyClient()	
  end if
  Set orderShip = nothing
  Set objShip = nothing

  if(Err.number <> 0) then 
  'response.write(Err.description)
  end if
end if	
%>