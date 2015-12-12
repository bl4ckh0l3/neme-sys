<%
Dim isAdmin, isEditor
isAdmin = false
isEditor = false

if not(isNull(request("cssClass"))) AND not(request("cssClass") = "") then
	cssClass = request("cssClass")
end if

if not(isEmpty(Session("objCMSUtenteLogged"))) then
	Dim objUsrLoggedMenu, objUsrLoggedMenuTmp, strRuolo
	Set objUsrLoggedMenuTmp = new UserClass
	Set objUsrLoggedMenu = objUsrLoggedMenuTmp.findUserByID(Session("objCMSUtenteLogged"))
	strRuolo = objUsrLoggedMenu.getRuolo()
	Set objUsrLoggedMenu = nothing
	Set objUsrLoggedMenuTmp = nothing

	if(strRuolo = Application("admin_role")) then
		isAdmin=true
	end if
	if(strRuolo = Application("editor_role")) then
		isEditor = true
	end if
end if
%>
<div id="backend-menu-left-container">
  <div id="backend-menu-left" style="<%if(Session("menu_closed")="1")then response.write("display:none;") end if%>">
    <%if(isAdmin OR isEditor) then%>		
    <ul>
      <li><a href="<%=Application("baseroot")&"/editor/contenuti/InserisciNews.asp?id_news=-1&cssClass=LN"%>"><img src="<%=Application("baseroot") & "/editor/img/page_white_edit.png"%>" border="0" title="<%=langEditor.getTranslated("backend.menu.item.contenuti.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/contenuti/ListaNews.asp?cssClass=LN&resetMenu=1"%>" class="<%if(strComp(cssClass, "LN", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.contenuti.lista")%>"><%=langEditor.getTranslated("backend.menu.item.contenuti")%></a></li>
    </ul>		
    <ul>
      <li><a href="<%=Application("baseroot")&"/editor/newsletter/InserisciNewsletter.asp?cssClass=LNL&id_newsletter=-1"%>"><img src="<%=Application("baseroot") & "/editor/img/newsletter_add.png"%>" border="0" title="<%=langEditor.getTranslated("backend.menu.item.newsletter.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/newsletter/ListaNewsletter.asp?cssClass=LNL"%>" class="<%if(strComp(cssClass, "LNL", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.newsletter.lista")%>"><%=langEditor.getTranslated("backend.menu.item.newsletter")%></a></li>
    </ul>			
    <%end if
    if(isAdmin) then%>
    <ul>		
      <li><a href="<%=Application("baseroot")&"/editor/utenti/InserisciUtente.asp?cssClass=LU&id_utente=-1"%>"><img src="<%=Application("baseroot") & "/editor/img/user_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.utenti.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/utenti/ListaUtenti.asp?cssClass=LU"%>" class="<%if(strComp(cssClass, "LU", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.utenti.lista")%>"><%=langEditor.getTranslated("backend.menu.item.utenti")%></a></li>
    </ul>		
    <ul>		
      <li><a href="<%=Application("baseroot")&"/editor/targets/InserisciTarget.asp?cssClass=LT&id_target=-1"%>"><img src="<%=Application("baseroot") & "/editor/img/link_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.target.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/targets/ListaTarget.asp?cssClass=LT"%>" class="<%if(strComp(cssClass, "LT", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.target.lista")%>"><%=langEditor.getTranslated("backend.menu.item.target")%></a></li>
    </ul>		
    <ul>		
      <li><a href="<%=Application("baseroot")&"/editor/categorie/InserisciCategoria.asp?cssClass=LCE&id_categoria=-1"%>"><img src="<%=Application("baseroot") & "/editor/img/folder_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.categorie.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/categorie/ListaCategorie.asp?cssClass=LCE"%>" class="<%if(strComp(cssClass, "LCE", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.categorie.lista")%>"><%=langEditor.getTranslated("backend.menu.item.categorie")%></a></li>
    </ul>		
    <ul>		
      <li><a href="<%=Application("baseroot")&"/editor/templates/InserisciTemplate.asp?cssClass=LTP&id_template=-1"%>"><img src="<%=Application("baseroot") & "/editor/img/layout_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.templates.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/templates/ListaTemplates.asp?cssClass=LTP"%>" class="<%if(strComp(cssClass, "LTP", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.templates.lista")%>"><%=langEditor.getTranslated("backend.menu.item.templates")%></a></li>
    </ul>	
    <ul>		
      <li><a href="<%=Application("baseroot")&"/editor/countries/InserisciCountry.asp?id_country=-1&cssClass=LCT"%>"><img src="<%=Application("baseroot") & "/editor/img/world_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.country.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/countries/ListaCountry.asp?cssClass=LCT&resetMenu=1"%>" class="<%if(strComp(cssClass, "LCT", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.country.lista")%>"><%=langEditor.getTranslated("backend.menu.item.country")%></a></li>
    </ul>		
    <ul>		
      <li><img src="<%=Application("baseroot") & "/editor/img/spacer.png"%>" border="0"/></li>
      <li><a href="<%=Application("baseroot")&"/editor/language/InserisciLingua.asp?cssClass=IL"%>" class="<%if(strComp(cssClass, "IL", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.language.lista")%>"><%=langEditor.getTranslated("backend.menu.item.language")%></a></li>
    </ul>			
    <%end if
    if(isAdmin OR isEditor) then%>
    <ul>		
      <li><img src="<%=Application("baseroot") & "/editor/img/spacer.png"%>" border="0"/></li>
      <li><a href="<%=Application("baseroot")&"/editor/multilanguage/InserisciMultiLingua.asp?cssClass=IML&resetMenu=1"%>" class="<%if(strComp(cssClass, "IML", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.multi_language.lista")%>"><%=langEditor.getTranslated("backend.menu.item.multi_language")%></a></li>
    </ul>			
    <%end if
    if(isAdmin) then%>
    <ul>				
      <li><img src="<%=Application("baseroot") & "/editor/img/spacer.png"%>" border="0"/></li>
      <li><a href="<%=Application("baseroot")&"/editor/configuration/PortalConfig.asp?cssClass=CP"%>" class="<%if(strComp(cssClass, "CP", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.config_portal.lista")%>"><%=langEditor.getTranslated("backend.menu.item.config_portal")%></a></li>
    </ul>		
    <ul>		
      <li><img src="<%=Application("baseroot") & "/editor/img/spacer.png"%>" border="0"/></li>
      <li><a href="<%=Application("baseroot")&"/editor/logs/ListaLogs.asp?cssClass=LL&resetMenu=1"%>" class="<%if(strComp(cssClass, "LL", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.logs_portal.lista")%>"><%=langEditor.getTranslated("backend.menu.item.logs_portal")%></a></li>
    </ul>		
    <%end if%>
    <br/>	
<!--nsys-editinc1-->
    <%if(isAdmin) then%>
    <ul>		
      <li><a href="<%=Application("baseroot")&"/editor/payments/InserisciPayment.asp?id_payment=-1&cssClass=LPT"%>"><img src="<%=Application("baseroot") & "/editor/img/money_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.payment.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/payments/ListaPayment.asp?cssClass=LPT"%>" class="<%if(strComp(cssClass, "LPT", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.payment.lista")%>"><%=langEditor.getTranslated("backend.menu.item.payment")%></a></li>
    </ul>		
    <%end if
    if(isAdmin OR isEditor) then%>
    <ul>		
      <li><a href="<%=Application("baseroot")&"/editor/currency/InserisciCurrency.asp?id_currency=-1&cssClass=LCY"%>"><img src="<%=Application("baseroot") & "/editor/img/money_euro.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.currency.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/currency/ListaCurrency.asp?cssClass=LCY"%>" class="<%if(strComp(cssClass, "LCY", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.currency.lista")%>"><%=langEditor.getTranslated("backend.menu.item.currency")%></a></li>
    </ul>		
    <ul>
      <li><a href="<%=Application("baseroot")&"/editor/tax/InserisciTassa.asp?id_tassa=-1&cssClass=LTX"%>"><img src="<%=Application("baseroot") & "/editor/img/coins_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.tasse.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/tax/ListaTasse.asp?cssClass=LTX"%>" class="<%if(strComp(cssClass, "LTX", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.tasse.lista")%>"><%=langEditor.getTranslated("backend.menu.item.tasse")%></a></li>
    </ul>		
    <ul>
      <li><a href="<%=Application("baseroot")&"/editor/spese/InserisciSpesaAccessoria.asp?id_spesa=-1&cssClass=LSP"%>"><img src="<%=Application("baseroot") & "/editor/img/lorry_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.spese.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/spese/ListaSpeseAccessorie.asp?cssClass=LSP"%>" class="<%if(strComp(cssClass, "LSP", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.spese.lista")%>"><%=langEditor.getTranslated("backend.menu.item.spese")%></a></li>
    </ul>		
    <ul>
      <li><a href="<%=Application("baseroot")&"/editor/prodotti/InserisciProdotto.asp?id_prodotto=-1&cssClass=LP"%>"><img src="<%=Application("baseroot") & "/editor/img/photo_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.prodotti.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/prodotti/ListaProdotti.asp?cssClass=LP&resetMenu=1"%>" class="<%if(strComp(cssClass, "LP", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.prodotti.lista")%>"><%=langEditor.getTranslated("backend.menu.item.prodotti")%></a></li>
    </ul>		
    <ul>
      <li><a href="<%=Application("baseroot")&"/editor/ordini/InserisciOrdine1.asp?id_ordine=-1&cssClass=LO"%>"><img src="<%=Application("baseroot") & "/editor/img/brick_add.png"%>" border="0" alt="<%=langEditor.getTranslated("backend.menu.item.ordini.inserisci")%>"/></a></li>
      <li><a href="<%=Application("baseroot")&"/editor/ordini/ListaOrdini.asp?cssClass=LO&resetMenu=1"%>" class="<%if(strComp(cssClass, "LO", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.ordini.lista")%>"><%=langEditor.getTranslated("backend.menu.item.ordini")%></a></li>
    </ul>		
    <ul>
      <li><img src="<%=Application("baseroot") & "/editor/img/spacer.png"%>" border="0"/></li>
      <li><a href="<%=Application("baseroot")&"/editor/carrelli/ListaCarrelli.asp?cssClass=LCI"%>" class="<%if(strComp(cssClass, "LCI", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.carrelli.lista")%>"><%=langEditor.getTranslated("backend.menu.item.carrelli")%></a></li>
    </ul>		
    <%end if
    if(isAdmin) then%>
    <ul>		
      <li><img src="<%=Application("baseroot") & "/editor/img/spacer.png"%>" border="0"/><!--<a href="<%'=Application("baseroot")&"/editor/margini/InserisciMargine.asp?id_margine=-1&cssClass=LM"%>"><img src="<%'=Application("baseroot") & "/editor/img/calculator_add.png"%>" border="0" alt="<%'=langEditor.getTranslated("backend.menu.item.margini.inserisci")%>"/></a>--></li>
      <li><a href="<%=Application("baseroot")&"/editor/margini/ListaMargini.asp?cssClass=LM&resetMenu=1"%>" class="<%if(strComp(cssClass, "LM", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.margini.lista")%>"><%=langEditor.getTranslated("backend.menu.item.margini")%></a></li>
    </ul>
<!--nsys-voucher1-->      
    <ul>		
      <li><img src="<%=Application("baseroot") & "/editor/img/spacer.png"%>" border="0"/></li>
      <li><a href="<%=Application("baseroot")&"/editor/voucher/ListaVoucher.asp?cssClass=LVC&resetMenu=1"%>" class="<%if(strComp(cssClass, "LVC", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.voucher.lista")%>"><%=langEditor.getTranslated("backend.menu.item.voucher")%></a></li>
    </ul>
<!---nsys-voucher1-->	
    <%end if%>
<!---nsys-editinc1-->
  </div>
  <div id="backend-menu-left-hide">
  <img src="<%=Application("baseroot") & "/editor/img/close_corner.png"%>" border="0" id="close_corner" align="absmiddle"/>
  </div>
</div>
<script>
function setCloseMenu(closed){
	var query_string = "menu_closed="+closed;

	$.ajax({
	   type: "POST",
	   url: "<%=Application("baseroot") & "/editor/include/menu_closed.asp"%>",
	   data: query_string,
		success: function() {
		}
	 });
}

$('#close_corner').click(function() {
	var element = document.getElementById("backend-menu-left");
	var closed; 
	if(element.style.display == 'none'){
		$('#backend-menu-left').show('slow');
		closed = 0;
	}else{
		$('#backend-menu-left').hide('slow');
		closed = 1;
	}
     
	setCloseMenu(closed);
});
</script>
