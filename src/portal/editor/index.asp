<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->

<%if (isEmpty(Session("objCMSUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp")
end if

Dim objUsrLoggedEditor, objUsrLoggedEditorTmp, strRuoloEditor
Set objUsrLoggedEditorTmp = new UserClass
Set objUsrLoggedEditor = objUsrLoggedEditorTmp.findUserByID(Session("objCMSUtenteLogged"))
strRuoloEditor = objUsrLoggedEditor.getRuolo()
Set objUsrLoggedEditor = nothing
Set objUsrLoggedEditorTmp = nothing%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<SCRIPT type=text/javascript>
/*$(document).ready(function(){
	$(".selectors a").hover(function() {
		$(this).next("em").animate({opacity: "show", top: "-30"}, "slow");
		$(this).children("span").animate({opacity: "show"}, "slow");
	}, function() {
		$(this).children("span").animate({opacity: "hide"}, "fast");
		$(this).next("em").animate({opacity: "hide", top: "-40"}, "fast");
	});
});*/


$(function() {
     $('img[data-hover]').hover(function() {
         $(this).attr('tmp', $(this).attr('src')).attr('src', $(this).attr('data-hover')).attr('data-hover', $(this).attr('tmp')).removeAttr('tmp');
     }).each(function() {
         $('<img />').attr('src', $(this).attr('data-hover'));
     });;
}); 


</SCRIPT>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<!-- #include virtual="/editor/include/menu.inc" -->		
		<div id="backend-content-home" align="center">
			<div align="center">
				<strong id="init"><br><br><%=langEditor.getTranslated("backend.index.detail.table.label.editor_contents")%><br><br></strong>
				<%if(isAdmin OR isEditor) then%>
				<a href="<%=Application("baseroot")&"/editor/contenuti/ListaNews.asp?cssClass=LN&resetMenu=1"%>" class="<%if(strComp(cssClass, "LN", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.contenuti.lista")%>"><IMG height="120" alt="<%=langEditor.getTranslated("backend.menu.item.contenuti")%>" src="<%=Application("baseroot")&"/editor/img/home/ico-contenuti.jpg"%>" data-hover="<%=Application("baseroot")&"/editor/img/home/ico-active-contenuti.jpg"%>" width="100" border="0" vspace="3" hspace="2"></a>
				<%end if
    			if(isAdmin) then%>
				<a href="<%=Application("baseroot")&"/editor/utenti/ListaUtenti.asp?cssClass=LU"%>" class="<%if(strComp(cssClass, "LU", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.utenti.lista")%>"><IMG height="120" alt="<%=langEditor.getTranslated("backend.menu.item.utenti")%>" src="<%=Application("baseroot")&"/editor/img/home/ico-utenti.jpg"%>" data-hover="<%=Application("baseroot")&"/editor/img/home/ico-active-utenti.jpg"%>" width="100" border="0" vspace="3"></a>
				<a href="<%=Application("baseroot")&"/editor/categorie/ListaCategorie.asp?cssClass=LCE"%>" class="<%if(strComp(cssClass, "LCE", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.categorie.lista")%>"><IMG height="120" alt="<%=langEditor.getTranslated("backend.menu.item.categorie")%>" src="<%=Application("baseroot")&"/editor/img/home/ico-struttura.jpg"%>" data-hover="<%=Application("baseroot")&"/editor/img/home/ico-active-struttura.jpg"%>" width="100" border="0" vspace="3" hspace="2"></a><br/>
				<a href="<%=Application("baseroot")&"/editor/templates/ListaTemplates.asp?cssClass=LTP"%>" class="<%if(strComp(cssClass, "LTP", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.templates.lista")%>"><IMG height="120" alt="<%=langEditor.getTranslated("backend.menu.item.templates")%>" src="<%=Application("baseroot")&"/editor/img/home/ico-grafica.jpg"%>" data-hover="<%=Application("baseroot")&"/editor/img/home/ico-active-grafica.jpg"%>" width="100" border="0" hspace="2"></a>
				<%end if
    			if(isAdmin OR isEditor) then%>
				<a href="<%=Application("baseroot")&"/editor/multilanguage/InserisciMultiLingua.asp?cssClass=IML&resetMenu=1"%>" class="<%if(strComp(cssClass, "IML", 1) = 0) then response.Write("active")%>" title="<%=langEditor.getTranslated("backend.menu.item.multi_language.lista")%>"><IMG height="120" alt="<%=langEditor.getTranslated("backend.menu.item.multi_language")%>" src="<%=Application("baseroot")&"/editor/img/home/ico-multilingua.jpg"%>" data-hover="<%=Application("baseroot")&"/editor/img/home/ico-active-multilingua.jpg"%>" width="100" border="0"></a><br/><br/>
				<%end if%>

				<div><strong><%=langEditor.getTranslated("backend.index.detail.table.label.download_guide")%></strong>&nbsp;<!--nsys-bohome1--><a class="link-down-guide" target="_blank" href="http://www.neme-sys.it/public/utils/econeme-sys_guide.pdf"><!---nsys-bohome1--><%=langEditor.getTranslated("backend.index.detail.table.label.download_guide_click")%></a></div>
			</div>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>