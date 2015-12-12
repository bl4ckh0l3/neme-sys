<%@Language=VBScript codepage=65001 %>
<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/common/include/Objects/CommentsClass.asp" -->
<!-- #include file="include/init3.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<SCRIPT SRC="<%=Application("baseroot") & "/common/js/hashtable.js"%>"></SCRIPT>
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="Javascript1.2">
var listPreviewGerContent;
listPreviewGerContent = new Hashtable();

<%
if(hasContentFields) then		
	jsRenderContentField = ""
	
	for each k in objListContentField
		Set objField = objListContentField(k)
		labelForm = objField.getDescription()
		if not(langEditor.getTranslated("backend.contenuti.detail.table.label."&objField.getDescription())="") then labelForm = langEditor.getTranslated("backend.contenuti.detail.table.label."&objField.getDescription())
		jsRenderContentField = jsRenderContentField & objContentField.renderContentFieldJS(objField,"form_inserisci","",langEditor,"",2)
		
		Set objField = nothing
	next
end if
%>
</SCRIPT>
<%
if (Cint(id_news) <> -1) then
	Dim objNews, objSelNews
	Set objNews = New NewsClass
	Set objSelNews = objNews.findNewsByID(id_news)
	Set objNews = nothing
	
	if not(Instr(1, typename(objSelNews), "NewsClass", 1) > 0) then
		response.Redirect(Application("baseroot")&Application("error_page")&"?error=004")
	end if
	
	id_news = objSelNews.getNewsID()
	strTitolo = objSelNews.getTitolo()
	strAbs1 = objSelNews.getAbstract1()
	strAbs2 = objSelNews.getAbstract2()
	strAbs3 = objSelNews.getAbstract3()
	strText = objSelNews.getTesto()	
	strKeyword = objSelNews.getKeyword()
	dtData_ins = objSelNews.getDataInsNews()
	dtData_pub = objSelNews.getDataPubNews()
	dtData_del = objSelNews.getDataDelNews()
	stato_news = objSelNews.getStato()
	page_title = objSelNews.getPageTitle()
	meta_description = objSelNews.getMetaDescription()
	meta_keyword = objSelNews.getMetaKeyword()
	
	if not(isNull(objSelNews.getListaTarget())) then
		Set objTarget = objSelNews.getListaTarget()
	end if	
		
	if not(isNull(objSelNews.getFilePerNews())) then
		Set objFiles = objSelNews.getFilePerNews()	
	end if
	
	if (Instr(1, typename(objTarget), "dictionary", 1) > 0) then%>
	<script language="Javascript1.2">
		<%dim catPreviewList, objTmpCatclass, objTmpPreviewList, objTmpPreview
		Set catPreviewList = Server.CreateObject("Scripting.Dictionary")	
		Set objTmpCatclass = new CategoryClass
		for each y in objTarget.Keys
			if (objTarget(y).getTargetType() = 1) then
				On Error Resume Next
				Set objTmpPreviewList = objTmpCatclass.findCategorieByTargetID(y)
				if not (isNull(objTmpPreviewList)) then
				for each j in objTmpPreviewList.Keys
					Set objTmpPreview = objTmpPreviewList(j)
					catPreviewList.add objTmpPreview.getCatGerarchia(), objTmpPreview.getCatDescrizione()
					Set objTmpPreview = nothing
				next
				end if
				Set objTmpPreviewList = nothing
				if(Err.number <>0)then
				end if
			end if
		next
		
		for each z in catPreviewList.Keys%>
			listPreviewGerContent.put("<%=z%>","<%=catPreviewList(z)%>");	
		<%next	
		Set objTmpCatclass = nothing
		Set catPreviewList = nothing%>
	</SCRIPT>
	<%end if						
end if

Dim objFilePerNews
Set objFilePerNews = new File4NewsClass
Set objListFileLabel = objFilePerNews.getListaFileLabel()
Set objFilePerNews = nothing
%>
<script language="Javascript1.2">
var templatePreviewFile = "";
var gerarchiaPreviewContent = "";
	
function preview(id_news){
	var templatePreviewPath = "";
	templatePreviewPath = templatePreviewPath + "<%=Application("baseroot")&Application("dir_upload_templ")%>" + "newsletter/<%=langEditor.getLangcode()%>/" + templatePreviewFile + "?id_news=" + id_news;
	if(templatePreviewFile != "" && id_news > 0){
		openWin(templatePreviewPath,'templatenewsletter',970,600,150,60);
	}else{
		alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.preview_disabled")%>");
	}
}
	
function previewContent(id_news){
	var templatePreviewPath = "";
	
	templatePreviewPath = templatePreviewPath + "<%=Application("baseroot")%>" + "/common/include/Controller.asp?is_preview_content=1&id_news=" + id_news+"&gerarchia="+gerarchiaPreviewContent;
	if(gerarchiaPreviewContent != "" && id_news > 0){
		openWin(templatePreviewPath,'templatecontent',970,600,150,60);
	}else{
		alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.preview_content_disabled")%>");
	}
}

function changeTemplatePreviewPath(templatePreviewFileChanged){
	templatePreviewFileChanged = templatePreviewFileChanged.substring(templatePreviewFileChanged.indexOf("|")+1);
	templatePreviewFile = templatePreviewFileChanged;
}

function changeTemplatePreviewContentGer(gerPreviewChanged){
	gerarchiaPreviewContent = gerPreviewChanged;
}


function changeNumMaxImgs(){
	if(document.form_inserisci.numMaxImgs.value == ""){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.insert_value")%>");
		document.form_inserisci.numMaxImgs.focus();
		return;
	}else if(isNaN(document.form_inserisci.numMaxImgs.value)){
		alert("<%=langEditor.getTranslated("backend.templates.detail.js.alert.isnan_value")%>");
		document.form_inserisci.numMaxImgs.focus();
		return;		
	}
	//location.href = "<%=Application("baseroot") & "/editor/contenuti/Inseriscinews.asp?id_news="&id_news&"&numMaxImgs="%>"+document.form_inserisci.numMaxImgs.value;
	renderNumImgsTable(document.form_inserisci.numMaxImgs.value);
}

function renderNumImgsTable(counter){
	$(".attach_table_rows").remove();
	
	var render ="";
	
	for(var i=1;i<=counter;i++){
		render=render+'<tr class="attach_table_rows">';
			render=render+'<td><input type="file" name="fileupload'+i+'" class="formFieldTXT"></td>';
			render=render+'<td>';
			render=render+'<select name="fileupload'+i+'_label" class="formFieldSelectTypeFile">';
			<%for each xType in objListFileLabel%>
			<%="render=render+'<option value="""&xType&""">"&objListFileLabel(xType)&"</option>';"%>
			<%next%>
			render=render+'</select>';
			render=render+'</td>';
			render=render+'<td><input type="text" name="fileupload'+i+'_dida" class="formFieldTXT"></td>';
			render=render+'<td>';
			if(i==1){
			render=render+'<input type="text" value="'+counter+'" name="numMaxImgs" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);"><a href="javascript:changeNumMaxImgs();"><img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="0" hspace="4" border="0" align="top" alt="<%=langEditor.getTranslated("backend.commons.detail.table.label.change_num_imgs")%>"></a>';
			}
			render=render+'</td>';
		render=render+'</tr>';
	}

	$("#add_attach_table").append(render);

}
</script>

<script language="JavaScript">
function sendForm(saveEsc){
	if(controllaCampiInput()){
	
		// se il controllo è andato a buon fine verifico se è stato selezionato il flag per l'inserimento file di grandi dimensioni,
		// in quel caso modifico la action per andare verso la pagina processnews2.asp, che utilizza la libreria ASPUPLOADLib.dll disponibile su aruba
		// verificare sempre che il provide rdisponga di quella libreria o che sia caricabile a parte, alrimenti eliminare l'opzione per la scelta del caricamento file di grandi dimensioni
		//if(document.form_inserisci.big_attachment && document.form_inserisci.big_attachment.checked){
			//document.form_inserisci.action="<%=Application("baseroot") & "/editor/contenuti/ProcessNews2.asp"%>";
		//}
		<%if(Application("use_aspupload_lib") = 1) then%>
			document.form_inserisci.action="<%=Application("baseroot") & "/editor/contenuti/ProcessNews2.asp"%>";
		<%end if%>
		
		document.form_inserisci.save_esc.value = saveEsc;
		document.getElementById("loading").style.visibility = "visible";
		document.getElementById("loading").style.display = "block";
		document.form_inserisci.submit();
	}else{
		return;
	}
}

function confirmDelete(){
	if(confirmDel()){
		document.form_cancella_news.submit();
	}else{
		return;
	}
}

function confirmDel(){
	return confirm('<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.confirm_del_news")%>');
}

function move(fbox, tbox){
	var arrFbox = new Array();
	var arrTbox = new Array();
	var arrLookup = new Array();
	var i;
	
	for(i = 0; i < tbox.options.length; i++){
		arrLookup[tbox.options[i].text] = tbox.options[i].value;
		arrTbox[i] = tbox.options[i].text;
	}
	
	var fLength = 0;
	var tLength = arrTbox.length;
	
	for(i = 0; i < fbox.options.length; i++){
		arrLookup[fbox.options[i].text] = fbox.options[i].value;
		if(fbox.options[i].selected && fbox.options[i].value != ""){
			arrTbox[tLength] = fbox.options[i].text;
			tLength++;
		}else{
			arrFbox[fLength] = fbox.options[i].text;
			fLength++;
		}
	}
	
	arrFbox.sort();
	arrTbox.sort();
	fbox.length = 0;
	tbox.length = 0;
	var c;
	
	for(c = 0; c < arrFbox.length; c++){
		var no = new Option();
		no.value = arrLookup[arrFbox[c]];
		no.text = arrFbox[c];
		fbox[c] = no;
	}
	
	for(c = 0; c < arrTbox.length; c++){
		var no = new Option();
		no.value = arrLookup[arrTbox[c]];
		no.text = arrTbox[c];
		tbox[c] = no;
	}
}

var dataInserimento = new Date();
var dataPubblicazione = new Date();

function controllaCampiInput(){

	/*
	 * codice per il controllo della data di pubblicazione: DA VERIFICARE
	 */
	
	if(!verificaData(document.form_inserisci.YEAR_PUBBLICAZIONE.value, document.form_inserisci.MONTH_PUBBLICAZIONE.value, document.form_inserisci.DAY_PUBBLICAZIONE.value, document.form_inserisci.HH_PUBBLICAZIONE.value, document.form_inserisci.MIN_PUBBLICAZIONE.value, "pub")){
		return false;
	}

	document.form_inserisci.news_data_pub.value = document.form_inserisci.YEAR_PUBBLICAZIONE.value + "-" + (parseInt(document.form_inserisci.MONTH_PUBBLICAZIONE.value)+1) + "-" + document.form_inserisci.DAY_PUBBLICAZIONE.value + " " + document.form_inserisci.HH_PUBBLICAZIONE.value + ":" + document.form_inserisci.MIN_PUBBLICAZIONE.value + ":59";
	
	/*
	 * codice per il controllo della data di cancellazione: DA VERIFICARE
	 */
	if(!parseInt(document.form_inserisci.YEAR_CANCELLAZIONE.value) == 0){
		if(!verificaData(document.form_inserisci.YEAR_CANCELLAZIONE.value, document.form_inserisci.MONTH_CANCELLAZIONE.value, document.form_inserisci.DAY_CANCELLAZIONE.value, document.form_inserisci.HH_CANCELLAZIONE.value, document.form_inserisci.MIN_CANCELLAZIONE.value, "del")){
			return false;
		}
		//*******************************************************************
		
		document.form_inserisci.news_data_del.value = document.form_inserisci.YEAR_CANCELLAZIONE.value + "-" + (parseInt(document.form_inserisci.MONTH_CANCELLAZIONE.value)+1) + "-" + document.form_inserisci.DAY_CANCELLAZIONE.value + " " + document.form_inserisci.HH_CANCELLAZIONE.value + ":" + document.form_inserisci.MIN_CANCELLAZIONE.value + ":00";
	}else{
		document.form_inserisci.news_data_del.value = "";
	}

	if (listTargetLang=="") {
		alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.insert_target")%>");
		return false;
	}	
	
	//valorizzo il campo nascosto "ListTarget" con la lista dei Target della news separati da "|"
	var strTargets = "";
	strTargets+=listTargetCat
	strTargets+=listTargetLang
	if(strTargets.charAt(strTargets.length -1) == "|"){
		strTargets = strTargets.substring(0, strTargets.length -1);
	}
	
	document.form_inserisci.ListTarget.value = strTargets;
	//alert(document.form_inserisci.ListTarget.value);



	if(document.form_inserisci.send_newsletter){
		if(document.form_inserisci.send_newsletter.value == 1 && document.form_inserisci.choose_newsletter.value == ""){
			alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.choose_newsletter_template")%>");
			document.form_inserisci.choose_newsletter.focus();
			return false;
		}else{
			var tmpValue = document.form_inserisci.choose_newsletter.value;
			tmpValue = tmpValue.substring(0, tmpValue.indexOf("|"));
			document.form_inserisci.choosenNewsletter.value=tmpValue;
		}
	}
	
	
	//recupero i valori dei checkbox con i file allegati
	//da eliminare
	var i;
	var strFiles = "";
	if(document.form_inserisci.existingFiles != null){
		if(document.form_inserisci.existingFiles.length == null){
			if(document.form_inserisci.existingFiles.checked){
				strFiles = strFiles + document.form_inserisci.existingFiles.value + "|";
			}
		}else{
			for(i=0; i<document.form_inserisci.existingFiles.length; i++){
				if(document.form_inserisci.existingFiles[i].checked){		
					strFiles = strFiles + document.form_inserisci.existingFiles[i].value + "|";
				}
			}
		}
	}
	if(strFiles.charAt(strFiles.length -1) == "|"){
		strFiles = strFiles.substring(0, strFiles.length -1);
	}	
	document.form_inserisci.ListFileDaEliminare.value = strFiles;

	
	if(document.form_inserisci.titolo.value == ""){
		alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.insert_title")%>");
		document.form_inserisci.titolo.focus();
		return false;
	}else{
		var strTitoloTmp = document.form_inserisci.titolo.value;
		//strTitoloTmp = replaceChars(strTitoloTmp);
		//strTitoloTmp = replaceChars2(strTitoloTmp);
		//strTitoloTmp = escape(strTitoloTmp);
		document.form_inserisci.titolo.value = strTitoloTmp;	
	}
		
	/*var strAbs1Tmp = document.form_inserisci.abstract1.value;
	strAbs1Tmp = replaceChars(strAbs1Tmp);
	strAbs1Tmp = replaceChars2(strAbs1Tmp);
	//strAbs1Tmp = replaceDefaultEditorChars(strAbs1Tmp);
	document.form_inserisci.abstract1.value = replaceChars(strAbs1Tmp);	

	var strAbs2Tmp = document.form_inserisci.abstract2.value;
	strAbs2Tmp = replaceChars(strAbs2Tmp);
	strAbs2Tmp = replaceChars2(strAbs2Tmp);
	//strAbs2Tmp = replaceDefaultEditorChars(strAbs2Tmp);
	document.form_inserisci.abstract2.value = replaceChars(strAbs2Tmp);	

	var strAbs3Tmp = document.form_inserisci.abstract3.value;
	strAbs3Tmp = replaceChars(strAbs3Tmp);
	strAbs3Tmp = replaceChars2(strAbs3Tmp);
	//strAbs3Tmp = replaceDefaultEditorChars(strAbs3Tmp);
	document.form_inserisci.abstract3.value = replaceChars(strAbs3Tmp);	

	var strTextTmp = document.form_inserisci.testo.value;
	strTextTmp = replaceChars(strTextTmp);
	strTextTmp = replaceChars2(strTextTmp);
	//strTextTmp = replaceDefaultEditorChars(strTextTmp);
	document.form_inserisci.testo.value = replaceChars(strTextTmp);*/	


	$("#add_attach_table").find("input:text[name*='_dida']").each(function(){
		$(this).attr('value', replaceChars($(this).val()));
		$(this).attr('value', replaceChars2($(this).val()));
	});

	$("#modify_attach_table").find("input:text[name*='fileDaModificare_']").each(function(){
		$(this).attr('value', replaceChars($(this).val()));
		$(this).attr('value', replaceChars2($(this).val()));
	});	
	

	<%
	if(hasContentFields) then		
		response.write(jsRenderContentField)	
	end if
	%>

	//recupero i valori dei checkbox con i field aggiuntivi da associare al prodotto
	var k;
	var strContentfields = "";
	if(document.form_inserisci.content_field_active != null){
		if(document.form_inserisci.content_field_active.length == null){
			if(document.form_inserisci.content_field_active.checked){
				strContentfields = strContentfields + document.form_inserisci.content_field_active.value + "|";
			}
		}else{
			for(k=0; k<document.form_inserisci.content_field_active.length; k++){
				if(document.form_inserisci.content_field_active[k].checked){		
					strContentfields = strContentfields + document.form_inserisci.content_field_active[k].value + "|";
				}
			}
		}
	}
	if(strContentfields.charAt(strContentfields.length -1) == "|"){
		strContentfields = strContentfields.substring(0, strContentfields.length -1);
	}

	document.form_inserisci.list_content_fields.value = strContentfields;

	return true;
}

function verificaData(yy, mm, dd, hh, minut, typeData){	
	if(yy.length < 4){
	  alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.insert_year_four")%>");
	  return false;		
	}else if (isNaN(yy)) {
	  alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.insert_year_number")%>");
	  return false;
	}else if (parseInt(yy) < 1900) {
	   alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.year_before_1900")%>");
	   return false;
	}else if(!checkDate(yy, mm, dd)){
	   alert("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.wrong_date")%>");
	   return false;		
	}else{
		if(typeData == "pub"){
			dataPubblicazione = new Date(parseInt(yy), parseInt(mm), parseInt(dd), parseInt(hh), parseInt(minut), 59);
		}
			
		return true;
	}
	
}

function confrontaDate(dataInserimento, dataPubblicazione){
	if(dataInserimento.getTime() <= dataPubblicazione.getTime()){
		return true;
	}else{
		return confirm('<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.pub_date_less_curr_date")%>');
	}
}


function checkDate(yy, mm, dd) {
	var myDayStr = dd
	var myMonthStr = mm
	var myYearStr = yy
	
	/* Using form values, create a new date object
	which looks like "Wed Jan 1 00:00:00 EST 1975". */
	var myDate = new Date(parseInt(myYearStr), parseInt(myMonthStr), parseInt(myDayStr));
	
	// Convert the date to a string so we can parse it.
	var myDate_month = myDate.getMonth();
	
	/* Split the string at every space and put the values into an array so,
	using the previous example, the first element in the array is "Wed", the
	second element is "Jan", the third element is "1", etc. */
	//var myDate_array = myDate_string.split( ' ' );
	
	/* If we entered "Feb 31, 1975" in the form, the "new Date()" function
	converts the value to "Mar 3, 1975". Therefore, we compare the month
	in the array with the month we entered into the form. If they match,
	then the date is valid, otherwise, the date is NOT valid. */
	if ( myDate_month != myMonthStr ) {
	  //alert( 'I\'m sorry, but "' + myDateStr + '" is NOT a valid date.' );
	  return false;
	} else {
	  //alert( 'Congratulations! "' + myDateStr + '" IS a valid date.' );
	  return true;
	} 
} 


function replaceChars(inString){
	var outString = inString;
	var pos= 0;

	// ricerca e escaping degli apici
	/*var quote2= -1;
	do {
		quote2= outString.indexOf('\'', pos);
		if (quote2 >= 0) {
			outString= outString.substring(0, quote2) + "&#39;" + outString.substring(quote2 +1);
			pos= quote2+2;
		}
	} while (quote2 >= 0);*/
	
	// ricerca e escaping dei new line
	pos= 0;
	var linefeed= -1;
	do {
		linefeed= outString.indexOf('\n', pos);
		if (linefeed >= 0) {
			outString= outString.substring(0, linefeed) + "\\n" + outString.substring(linefeed +1);
			pos= linefeed+2;
		}
	} while (linefeed >= 0);
	
	// ricerca e escaping dei line feed
	pos= 0;
	var creturn= -1;
	do {
		creturn= outString.indexOf('\r', pos);
		if (creturn >= 0) {
			outString= outString.substring(0, creturn) + "\\r" + outString.substring(creturn +1);
			pos= creturn+2;
		}
	} while (creturn >= 0);

	//ricerca lettere accentate èéàòùì
	//&egrave;&eacute;&agrave;&ograve;&ugrave;&igrave;
	pos= 0;
	var letter= -1;
	do {
		letter= outString.indexOf('è', pos);
		if (letter >= 0) {
			outString= outString.substring(0, letter) + "&egrave;" + outString.substring(letter +1);
			pos= letter+2;
		}
	} while (letter >= 0);
	letter= -1;
	do {
		letter= outString.indexOf('é', pos);
		if (letter >= 0) {
			outString= outString.substring(0, letter) + "&eacute;" + outString.substring(letter +1);
			pos= letter+2;
		}
	} while (letter >= 0);
	letter= -1;
	do {
		letter= outString.indexOf('à', pos);
		if (letter >= 0) {
			outString= outString.substring(0, letter) + "&agrave;" + outString.substring(letter +1);
			pos= letter+2;
		}
	} while (letter >= 0);
	letter= -1;
	do {
		letter= outString.indexOf('ò', pos);
		if (letter >= 0) {
			outString= outString.substring(0, letter) + "&ograve;" + outString.substring(letter +1);
			pos= letter+2;
		}
	} while (letter >= 0);
	letter= -1;
	do {
		letter= outString.indexOf('ù', pos);
		if (letter >= 0) {
			outString= outString.substring(0, letter) + "&ugrave;" + outString.substring(letter +1);
			pos= letter+2;
		}
	} while (letter >= 0);
	letter= -1;
	do {
		letter= outString.indexOf('ì', pos);
		if (letter >= 0) {
			outString= outString.substring(0, letter) + "&igrave;" + outString.substring(letter +1);
			pos= letter+2;
		}
	} while (letter >= 0);
	
	// ricerca degli href e delle ancore	
	pos= 0;
	var href= -1;
	do {
		href= outString.indexOf('href=', pos);
		if (href >= 0) {
			var url = outString.substring(href,outString.indexOf('>', href));
			if(url.indexOf('#') >=0){
				outString= outString.substring(0, href+6) + outString.substring(outString.indexOf('#', href+6));
			}
			pos= href+6;
		}
	} while (href >= 0);
	

	return outString;	
}

function replaceChars2(inString){
	var outString2 = inString;
	var pos2= 0;

	// ricerca e escaping degli apici
	var quote2= -1;
	do {
		quote2= outString2.indexOf('\'', pos2);
		if (quote2 >= 0) {
			outString2= outString2.substring(0, quote2) + "&#39;" + outString2.substring(quote2 +1);
			pos2= quote2+2;
		}
	} while (quote2 >= 0);
	
	// ricerca e escaping dei doppi apici
	pos2= 0;
	var doublequote= -1;
	do {
		doublequote= outString2.indexOf('\"', pos2);
		if (doublequote >= 0) {
			outString2= outString2.substring(0, doublequote) + "&quot;" + outString2.substring(doublequote +1);
			pos2= doublequote+2;
		}
	} while (doublequote >= 0);

	return outString2;	
}

function replaceDefaultEditorChars(inString){
	var outStringDef = inString;

	// ricerca caratteri di default dell'editor html: <br type=&quot;_moz&quot; /> oppure <br type="_moz" /> oppure &lt;br type=&quot;_moz&quot; /&gt; oppure &lt;br /&gt; oppure <br />
	if(outStringDef =='<br type=&quot;_moz&quot; />' || outStringDef =='<br type="_moz" />' || outStringDef =='&lt;br type=&quot;_moz&quot; /&gt;' || outStringDef =='&lt;br /&gt;' || outStringDef =='<br />'){
		outStringDef = "";
	}

	return outStringDef;	
}

function showHideDivArrow(elemDiv,elemArrow){
	var elementDiv = document.getElementById(elemDiv);
	var elementArrow = document.getElementById(elemArrow);
	if(elementDiv.style.visibility == 'visible'){
		elementArrow.src='<%=Application("baseroot")&"/editor/img/div_freccia.gif"%>';
	}else if(elementDiv.style.visibility == 'hidden'){
		elementArrow.src='<%=Application("baseroot")&"/editor/img/div_freccia2.gif"%>';
	}
}

function showAllContentField(){
	var ischecked;
	
	if($('#activate_all_content_field:checked').val() == undefined){
		$('#inner-table-content-field-list tbody tr').each( function(){
			var objParent = $(this).find('input:checkbox[name="content_field_active"]');
			var obj = $(this).find('input:checkbox[name="content_field_active"]:checked');
			//alert(objParent.length);
			if(objParent.length > 0){
				//alert(obj.length);
				if(obj.length == 0){
					$(this).hide();
				}
			}
		});	
	}else{
		$('#inner-table-content-field-list tbody tr').each( function(){		
			$(this).show();
		});	
	}
}
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="IN"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
			<!-- #include virtual="/fckeditor/fckeditor.asp" -->	
			<table border="0" cellspacing="0" cellpadding="0" class="principal">
			<tr> 		  		  
				<td>
					<form action="<%=Application("baseroot") & "/editor/contenuti/ProcessNews.asp"%>" method="post" name="form_inserisci" enctype="multipart/form-data" accept-charset="UTF-8">
					<input type="hidden" value="<%=id_news%>" name="id_news">
					<input type="hidden" value="<%=dtData_ins%>" name="news_data">
					<input type="hidden" value="" name="news_data_pub">
					<input type="hidden" value="" name="news_data_del">
					<input type="hidden" value="1" name="save_esc">				  
					<input type="hidden" value="" name="choosenNewsletter">
					<input type="hidden" value="<%="http://" & request.ServerVariables("SERVER_NAME") %>" name="srv_name">
					<%
					'*************** INIZIALIZZO IL CODICE PER GENERARE GLI EDITOR HTML
					Dim oFCKeditor
					Set oFCKeditor = New FCKeditor
					'oFCKeditor.Width = 200
					oFCKeditor.Height = 200
					oFCKeditor.BasePath = "/fckeditor/"
					%>	  
					
					<div class="labelForm" align="left"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.title")%></div>
					<div id="divTitle" align="left">
					<textarea name="titolo" class="formFieldTXTAREAAbstract"><%if(not isNull(strTitolo) AND Trim(strTitolo)<>"")then response.write(Server.HTMLEncode(strTitolo)) end if%></textarea>
					</div><br/>
					<div align="left" style="float:left;padding-right: 5px;">				
						<span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.page_title")%></span><br/>
						<input type="text" name="page_title" value="<%if(not isNull(page_title) AND Trim(page_title)<>"")then response.write(Server.HTMLEncode(Trim(page_title))) end if%>" class="formFieldTXT">
					  </div>
					  <div align="left" style="float:left;padding-right: 5px;">
					  <span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.meta_description")%></span><br/>
					    <input type="text" name="meta_description" value="<%if(not isNull(meta_description) AND Trim(meta_description)<>"")then response.write(Server.HTMLEncode(Trim(meta_description))) end if%>" class="formFieldTXT">
					  </div>
					 <div align="left" style="padding-bottom:20px;">
					 <span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.meta_keyword")%></span><br/>
					    <input type="text" name="meta_keyword" value="<%if(not isNull(meta_keyword) AND Trim(meta_keyword)<>"")then response.write(Server.HTMLEncode(Trim(meta_keyword))) end if%>" class="formFieldTXT">
					</div>
					<br>
					
					<div class="divDetailHeader" align="left" onClick="javascript:showHideDiv('divSummary1');showHideDivArrow('divSummary1','arrow1');"><img src="<%if not(strAbs1 = "")then response.Write(Application("baseroot")&"/editor/img/div_freccia.gif") else response.Write(Application("baseroot")&"/editor/img/div_freccia2.gif") end if%>" vspace="0" hspace="0" border="0" align="right" id="arrow1"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.abstract_field")%></div>
					<div id="divSummary1"  <%if not(strAbs1 = "")then response.Write("style=""visibility:visible;display:block;""") else response.Write("style=""visibility:hidden;display:none;""") end if%> align="left">
					<%
					oFCKeditor.Value = strAbs1
					oFCKeditor.Create "abstract1"
					%>
					</div><br>
					 
					<div class="divDetailHeader" align="left" onClick="javascript:showHideDiv('divSummary2');showHideDivArrow('divSummary2','arrow2');"><img src="<%if not(strAbs2 = "")then response.Write(Application("baseroot")&"/editor/img/div_freccia.gif") else response.Write(Application("baseroot")&"/editor/img/div_freccia2.gif") end if%>" vspace="0" hspace="0" border="0" align="right" id="arrow2"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.abstract_field")%></div>
					<div id="divSummary2"  <%if not(strAbs2 = "")then response.Write("style=""visibility:visible;display:block;""") else response.Write("style=""visibility:hidden;display:none;""") end if%> align="left">
					<%
					oFCKeditor.Value = strAbs2
					oFCKeditor.Create "abstract2"
					%>
					</div><br>
					  
					<div class="divDetailHeader" align="left" onClick="javascript:showHideDiv('divSummary3');showHideDivArrow('divSummary3','arrow3');"><img src="<%if not(strAbs3 = "")then response.Write(Application("baseroot")&"/editor/img/div_freccia.gif") else response.Write(Application("baseroot")&"/editor/img/div_freccia2.gif") end if%>" vspace="0" hspace="0" border="0" align="right" id="arrow3"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.abstract_field")%></div>
					<div id="divSummary3"  <%if not(strAbs3 = "")then response.Write("style=""visibility:visible;display:block;""") else response.Write("style=""visibility:hidden;display:none;""") end if%> align="left">
					<%
					oFCKeditor.Value = strAbs3
					oFCKeditor.Create "abstract3"
					%>
					</div><br>
					
					<div class="divDetailHeader" align="left" onClick="javascript:showHideDiv('divText');showHideDivArrow('divText','arrowText');"><img src="<%if not(strText = "")then response.Write(Application("baseroot")&"/editor/img/div_freccia.gif") else response.Write(Application("baseroot")&"/editor/img/div_freccia2.gif") end if%>" vspace="0" hspace="0" border="0" align="right" id="arrowText"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.text")%></div>
					<div id="divText"  <%if not(strText = "")then response.Write("style=""visibility:visible;display:block;""") else response.Write("style=""visibility:hidden;display:none;""") end if%> align="left">
					<%
					oFCKeditor.Height = 400
					oFCKeditor.Value = strText
					oFCKeditor.Create "testo"
					%>	
					</div><br>

					<div class="labelForm" align="left"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.keyword")%></div>
					<div id="divTitle" align="left">
					<input type="text" name="keyword" value="<%if(not isNull(strKeyword) AND Trim(strKeyword)<>"")then response.write(Server.HTMLEncode(strKeyword)) end if%>" class="formFieldTXTLong" />
					</div><br>
					
					<%
					'***** inizializzo gli elementi per la googlemap
					strID=id_news					
					strType=1
					
					if(Cint(strID)=-1)then
						Set objGUID = new GUIDClass
						strID=objGUID.CreateNumberGUIDRandomVarLenght(7)*(-1)
						Set objGUID = nothing%>
						<input type="hidden" value="<%=strID%>" name="pregeoloc_el_id">
					<%end if%>
					<!-- #include virtual="/editor/include/localization_widget.asp" -->

					<input type="hidden" value="" name="ListTarget">					
					<%
					Set objT = New TargetClass
					response.write(objT.renderTargetBox("listTargetCat", "targetcatbox_sx","targetcatbox_dx",langEditor.getTranslated("backend.contenuti.detail.table.label.target_x_contenuti_cat"), langEditor.getTranslated("backend.contenuti.detail.table.label.target_disp_cat"), "1", objTarget, objListaTargetPerUser, false, false, langEditor))
					Set objT = Nothing
					%>
					<br/><br/>						
					<%
					Set objT = New TargetClass
					response.write(objT.renderTargetBox("listTargetLang", "targetlangbox_sx","targetlangbox_dx",langEditor.getTranslated("backend.contenuti.detail.table.label.target_x_contenuti_lang"), langEditor.getTranslated("backend.contenuti.detail.table.label.target_disp_lang"), "3", objTarget, objListaTargetPerUser, false, false, langEditor))
					Set objT = Nothing
					%>						

					<br>
					<div class="divDetailHeader" align="left" onClick="javascript:showHideDiv('divAttachments');showHideDivArrow('divAttachments','arrowAttach');"><img src="<%if (Instr(1, typename(objFiles), "Dictionary", 1) > 0)then response.Write(Application("baseroot")&"/editor/img/div_freccia.gif") else response.Write(Application("baseroot")&"/editor/img/div_freccia2.gif") end if%>" vspace="0" hspace="0" border="0" align="right" id="arrowAttach"><%=langEditor.getTranslated("backend.news.view.table.label.attached_files")%></div>
					<div id="divAttachments" <%if (Instr(1, typename(objFiles), "Dictionary", 1) > 0) then response.Write("style=""visibility:visible;display:block;""") else response.Write("style=""visibility:hidden;display:none;""") end if%> align="left">
					<table border="0" cellspacing="0" cellpadding="0" class="principal" id="add_attach_table">
					  <tr>
						<td><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.attachment")%></span></td>
						<td><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.file_type_label")%></span></td>
						<td><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.file_dida")%></span></td>
						<td><span class="labelForm"><%=langEditor.getTranslated("backend.commons.detail.table.label.change_num_imgs")%></span></td>
					  </tr>
					  <%
					  Dim fileCounter
					  for fileCounter=1 to numMaxImg%>
					  <tr class="attach_table_rows">
						<td><input type="file" name="fileupload<%=fileCounter%>" class="formFieldTXT"></td>
						<td>
						<select name="fileupload<%=fileCounter%>_label" class="formFieldSelectTypeFile">
						<%for each xType in objListFileLabel%>
						<option value="<%=xType%>"><%=objListFileLabel(xType)%></option>
						<%next%>
						</select>
						</td>
						<td><input type="text" name="fileupload<%=fileCounter%>_dida" class="formFieldTXT"></td>
						<td><%if(fileCounter=1)then%><input type="text" value="<%=numMaxImg%>" name="numMaxImgs" class="formFieldTXTShort" onkeypress="javascript:return isInteger(event);"><a href="javascript:changeNumMaxImgs();"><img src=<%=Application("baseroot")&"/common/img/refresh.gif"%> vspace="0" hspace="4" border="0" align="top" alt="<%=langEditor.getTranslated("backend.commons.detail.table.label.change_num_imgs")%>"></a><%end if%>&nbsp;</td>
					  </tr>
					 <%next%>	 
					</table>
				  
					<%Dim listFileToModify
					listFileToModify = ""
					if not(isNull(objFiles)) then%>		
						<table border="0" cellspacing="0" cellpadding="0" class="principal" id="modify_attach_table">
						  <tr>
							<td><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.file_to_del")%></span></td>
							<td><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.file_type_label")%></span></td>
							<td><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.file_dida")%></span></td>
						  </tr>
						<%Dim objFilesInNews
						for each z in objFiles.Keys
							Set objFilesInNews = objFiles(z)%>
						  <tr>
							<td><input type="checkbox" value="<%=objFilesInNews.getFileID()%>" name="existingFiles">&nbsp;<%=objFilesInNews.getFileName()%></td>
							<td>
							<select name="fileDaModificare_<%=objFilesInNews.getFileID()%>_label" class="formFieldSelectTypeFile">
							<%for each xType in objListFileLabel%>
							<option value="<%=xType%>" <%if(xType=objFilesInNews.getFileTypeLabel()) then response.write("selected") end if%>><%=objListFileLabel(xType)%></option>
							<%next%>
							</select>
							</td>
							<td><input type="text" name="fileDaModificare_<%=objFilesInNews.getFileID()%>" value="<%=objFilesInNews.getFileDida()%>" class="formFieldTXT"></td>
						  </tr>
							<%listFileToModify = listFileToModify & objFilesInNews.getFileID() & "|"
							Set objFilesInNews = nothing	
						next				
						Set objFiles = nothing
						listFileToModify = Mid(listFileToModify, 1, (Len(listFileToModify)-1))%>
						</table>
					<%end if%>
					<input type="hidden" value="<%=Trim(listFileToModify)%>" name="ListFileDaModificare">
					<input type="hidden" value="" name="ListFileDaEliminare">				
					</div>
					
					<div style="float:left;padding-right:40px;padding-top:20px;">
					  <!-- ********************************** CAMPI PER DATA PUBBLICAZIONE ************************* -->
						<span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.data_pub")%></span><br>
						  <%
						  Dim RANGE_MINUTI, DD, MM, YY, HH, MIN, DATA_CORRENTE
						  DATA_CORRENTE = NOW()
						  RANGE_MINUTI = 1
						  if not(dtData_pub = "") then
							  DD = DatePart("d", dtData_pub)
							  MM = DatePart("m", dtData_pub)
							  YY = DatePart("yyyy", dtData_pub)
							  HH = DatePart("h", dtData_pub)
							  MIN = DatePart("n", dtData_pub)
						  else
							  DD = DatePart("d", DATA_CORRENTE)
							  MM = DatePart("m", DATA_CORRENTE)
							  YY = DatePart("yyyy", DATA_CORRENTE)
							  HH = DatePart("h", DATA_CORRENTE)
							  MIN = DatePart("n", DATA_CORRENTE)	  	
						  end if%>
						  <SELECT NAME="DAY_PUBBLICAZIONE" class="formFieldSelectDataPub">
							<%for i=1 to 31
							   if(DD = i) then%>
							<OPTION SELECTED VALUE="<%=i%>"><%=i%></OPTION>
							<%else%>
							<OPTION VALUE="<%=i%>"><%=i%></OPTION>
							<%end if
							next%>
						  </SELECT>
						  <SELECT NAME="MONTH_PUBBLICAZIONE" class="formFieldSelectDataPub">
						<%for i=0 to 11
							if(MM = i+1) then%>
								<OPTION SELECTED VALUE="<%=i%>"><%=i+1%></OPTION>
							<%else%>
								<OPTION VALUE="<%=i%>"><%=i+1%></OPTION>
							<%end if
						next%>
						  </SELECT>
						  <INPUT NAME="YEAR_PUBBLICAZIONE" TYPE="TEXT" VALUE="<%=YY%>" maxlength="4" class="formFieldAnnoPubblicazione">
						  -
						  <SELECT NAME="HH_PUBBLICAZIONE" class="formFieldSelectDataPub">
						<%for i=0to 23
							if(HH = i) then%>
								<OPTION SELECTED VALUE="<%=i%>"><%=i%></OPTION>
							<%else%>
								<OPTION VALUE="<%=i%>"><%=i%></OPTION>
							<%end if
						next%>
						  </SELECT>
						  :
						  <SELECT NAME="MIN_PUBBLICAZIONE" class="formFieldSelectDataPub">
						<%for i=0 to 59
							if(MIN = i) then%>
								<OPTION SELECTED VALUE="<%=i*RANGE_MINUTI%>"><%=i*RANGE_MINUTI%></OPTION>
							<%else%>
								<OPTION VALUE="<%=i*RANGE_MINUTI%>"><%=i*RANGE_MINUTI%></OPTION>
							<%end if
						next%>
						  </SELECT>
						</div>
						<div style="float:top;padding-top:20px;">	  
						<!-- ********************************** CAMPI PER DATA CANCELLAZIONE ************************* -->
						<span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.data_del")%></span><br>
						  <%
						  Dim RANGE_MINUTI_del, DD_del, MM_del, YY_del, HH_del, MIN_del			  
						  RANGE_MINUTI_del = 1
						  if not(dtData_del = "") then
							  DD_del = DatePart("d", dtData_del)
							  MM_del = DatePart("m", dtData_del)
							  YY_del = DatePart("yyyy", dtData_del)
							  HH_del = DatePart("h", dtData_del)
							  MIN_del = DatePart("n", dtData_del)
						  else
							  DD_del = 0
							  MM_del = 0
							  YY_del = 0
							  HH_del = 0
							  MIN_del = 0	  	
						  end if%>
						  <SELECT NAME="DAY_CANCELLAZIONE" class="formFieldSelectDataPub">
							<%for i=0 to 31
							   if(DD_del = i) then%>
							<OPTION SELECTED VALUE="<%=i%>"><%=i%></OPTION>
							<%else%>
							<OPTION VALUE="<%=i%>"><%=i%></OPTION>
							<%end if
							next%>
						  </SELECT>
						  <SELECT NAME="MONTH_CANCELLAZIONE" class="formFieldSelectDataPub">
							<%if(MM_del = 0) then%>
								<OPTION SELECTED VALUE="0">0</OPTION>
								<%for i=0 to 11%>
									<OPTION VALUE="<%=i%>"><%=i+1%></OPTION>
								<%next	
							else%>
								<OPTION VALUE="0">0</OPTION>				
								<%for i=0 to 11
									if(MM_del = i+1) then%>
										<OPTION SELECTED VALUE="<%=i%>"><%=i+1%></OPTION>
									<%else%>
										<OPTION VALUE="<%=i%>"><%=i+1%></OPTION>
									<%end if
								next
							end if%>
						  </SELECT>
						  <INPUT NAME="YEAR_CANCELLAZIONE" TYPE="TEXT" VALUE="<%=YY_del%>" maxlength="4" class="formFieldAnnoPubblicazione">
						  -
						  <SELECT NAME="HH_CANCELLAZIONE" class="formFieldSelectDataPub">
						<%for i=0 to 23
							if(HH_del = i) then%>
								<OPTION SELECTED VALUE="<%=i%>"><%=i%></OPTION>
							<%else%>
								<OPTION VALUE="<%=i%>"><%=i%></OPTION>
							<%end if
						next%>
						  </SELECT>
						  :
						  <SELECT NAME="MIN_CANCELLAZIONE" class="formFieldSelectDataPub">
						<%for i=0 to 59
							if(MIN_del = i) then%>
								<OPTION SELECTED VALUE="<%=i*RANGE_MINUTI_del%>"><%=i*RANGE_MINUTI_del%></OPTION>
							<%else%>
								<OPTION VALUE="<%=i*RANGE_MINUTI_del%>"><%=i*RANGE_MINUTI_del%></OPTION>
							<%end if
						next%>
						  </SELECT>
						</div>  
					
					<br><br>
					<div style="float:left;"><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.stato_contenuto")%></span><br>
					<select name="stato_news" class="formFieldTXT">
					<option value="0" <%if (stato_news = 0) then response.Write("selected")%>><%=langEditor.getTranslated("backend.contenuti.lista.table.select.option.edit")%></option>
					<option value="1" <%if (stato_news = 1) then response.Write("selected")%>><%=langEditor.getTranslated("backend.contenuti.lista.table.select.option.public")%></option>
					</select>&nbsp;&nbsp;</div>
					<div>
					<span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.button.label.preview_contenuti")%></span><br>
					<%if (Cint(id_news) <> -1) then%>
						<select name="choose_preview_cat" class="formFieldTXT" onChange="changeTemplatePreviewContentGer(this.value);">
						<option value=""></option>
						<script language="Javascript1.2">
						var arrKeys = listPreviewGerContent.keys();
						var tmpKey;
						var tmpValue;
						
						for(var z=0; z<arrKeys.length; z++){
							tmpKey = arrKeys[z];
							tmpValue = listPreviewGerContent.get(tmpKey);			
							document.write("<option value=\""+tmpKey+"\">"+tmpValue+"</option>");
						}
						</SCRIPT>		
						</select>	
						<a href="javascript:previewContent('<%=id_news%>')"><%=langEditor.getTranslated("backend.contenuti.detail.button.label.preview_contenuti")%></a>
						<br>
					<%end if%>
					</div><br><br>				
					
					<%Set objNewsletter = new NewsletterClass			
					Dim hasNewsletter, objNewsletterTmp
					hasNewsletter = false
					on error Resume Next
					
					Set objListaNewsletter = objNewsletter.getListaNewsletter(1)
					if isObject(objListaNewsletter) AND not(isNull(objListaNewsletter)) AND not (isEmpty(objListaNewsletter)) then
						if(objListaNewsletter.Count > 0) then
							hasNewsletter = true
						end if
					end if
					
					if Err.number <> 0 then
					end if
					
					if(hasNewsletter) then%>	  
						<span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.send_as_newsletter")%></span><br>
						<select name="send_newsletter" class="formFieldSelectNewsletter">
						<option value="0" selected><%=langEditor.getTranslated("backend.commons.no")%></option>
						<option value="1"><%=langEditor.getTranslated("backend.commons.yes")%></option>
						</select>&nbsp;
						<select name="choose_newsletter" class="formFieldTXT" onChange="changeTemplatePreviewPath(this.value);">
						<option value=""></option>
						<%for each x in objListaNewsletter.Keys			
							Set objNewsletterTmp = objListaNewsletter(x)%>
							<option value="<%=x&"|"&objNewsletterTmp.getTemplate()%>"><%=objNewsletterTmp.getDescrizione()%></option>
							<%Set objNewsletterTmp = nothing
						next%>
						</select>	
					<%end if%>
					<%Set objNewsletter = nothing%>
					<a href="javascript:preview('<%=id_news%>')"><%=langEditor.getTranslated("backend.contenuti.detail.button.label.preview_newsletter")%></a>  

					<br/><br/>
					<div class="divDetailHeader" align="left" onClick="javascript:showHideDiv('divContentFields');showHideDivArrow('divContentFields','arrowFields');"><img src="<%if (hasContentFields)then response.Write(Application("baseroot")&"/editor/img/div_freccia.gif") else response.Write(Application("baseroot")&"/editor/img/div_freccia2.gif") end if%>" vspace="0" hspace="0" border="0" align="right" id="arrowFields"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.product_fields")%></div>
					<div id="divContentFields" <%if (hasContentFields) then response.Write("style=""visibility:visible;display:block;padding-top:2px;""") else response.Write("style=""visibility:hidden;display:none;padding-top:2px;""") end if%> align="left">
					  <input type="hidden" value="" name="list_content_fields">
					  <input type="hidden" value="" name="list_content_fields_values">
					  <table border="0" align="top" cellpadding="0" cellspacing="0" class="inner-table" id="inner-table-content-field-list">
						<tr>
						<th width="155"><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.prod_field_attivo")%></span>&nbsp;&nbsp;<span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.prod_field_attivo_all")%></span><input type="checkbox" id="activate_all_content_field" value="" onclick="javascript:showAllContentField();" checked></th>
						<th width="20%"><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.prod_field_name")%></span></th>
						<th colspan="2"><span class="labelForm"><%=langEditor.getTranslated("backend.contenuti.detail.table.label.prod_field_values")%></span></th>
						</tr>
					  <tbody>
					<%
					Dim fieldCssClass
					if(hasContentFields) then
						Dim styleRow, styleRow2, counter
						styleRow2 = "table-list-on"
						counter = 0
					
						for each k in objListContentField
							styleRow = "table-list-off"
							if(counter MOD 2 = 0) then styleRow = styleRow2 end if
							
							Set objField = objListContentField(k)
							fieldCssClass=""
							
							if(CInt(objField.getTypeField())=4) then
								fieldCssClass="formFieldMultiple"
							end if
							
							labelForm = objField.getDescription()
							if not(langEditor.getTranslated("backend.contenuti.detail.table.label."&objField.getDescription())="") then labelForm = langEditor.getTranslated("backend.contenuti.detail.table.label."&objField.getDescription())
							%>		  
							<tr class="<%=styleRow%>" id="tr_content_field_<%=objField.getID()%>">
							<td id="td_content_field_active_<%=objField.getID()%>">
							<input type="checkbox" value="<%=objField.getID()&"-"&objField.getTypeField()%>" id="content_field_active_<%=objField.getID()&"-"&objField.getTypeField()%>" name="content_field_active" <%if(objField.getidContent() <> "")then response.write("checked='checked'") end if%>>&nbsp;</td>
							<td id="td_content_field_name_<%=objField.getID()%>"><%=labelForm%>&nbsp;</td>
							<td colspan="2"><%								
								select Case objField.getTypeField()
								'Case 3,4,5,6						
								'	On Error Resume next
								'	hasListValues = false
									
								'	Set objListValues = objContentField.getListContentFieldValues(k)
								'	if(objListValues.Count > 0)then
								'		hasListValues = true
								'	end if
								
								'	if(Err.number<>0) then
								'		'response.write(Err.description)
								'		hasListValues = false
								'	end if
								
								'	if(hasListValues)then
								'		Dim valueList
								'		valueList = ""
								'		for each g in objListValues
								'			valueList = valueList & Server.HTMLEncode(g) & ","
								'		next
										
								'		valueList = Left(valueList,InStrRev(valueList,",",-1,1)-1)						
								'		response.write(valueList)
										
								'		Set objListValues = nothing
								'	end if
								case 7
									fieldValueMatch = objContentField.findFieldMatchValue(k,id_news)
									response.write(objContentField.renderContentFieldHTML(objField,fieldCssClass, "", id_news, fieldValueMatch,langEditor,0,objField.getEditable()))%>
									<script>
									document.getElementById('<%=objContentField.getFieldPrefix()&objField.getID()%>').setAttribute('type', 'text');
									</script>
								<%case 8
								case else
									fieldValueMatch = objContentField.findFieldMatchValue(k,id_news)									
									response.write(objContentField.renderContentFieldHTML(objField,fieldCssClass, "", id_news, fieldValueMatch,langEditor,0,objField.getEditable()))						
								end select%></td>					
							 </tr>
						<%counter = counter +1
						next
					end if

					Set objListContentField = nothing
					Set objContentField = nothing
					%>
						</tbody>
					  </table>
					  </div>



					<%if (Cint(id_news) <> -1) then%>
						<br><br><br>
						<span class="labelForm"><%=langEditor.getTranslated("backend.news.view.table.label.comments")%></span><br>
						<%
						Set objCommento = New CommentsClass
						if(not(isNull(objCommento.findCommentiByIDElement(id_news,1,null)))) then%>
						<a href="javascript:openWin('<%=Application("baseroot")&"/public/layout/include/popupInsertComments.asp?id_element="&id_news&"&element_type=1"%>','popupallegati',400,400,100,100);" title="<%=langEditor.getTranslated("backend.news.view.table.label.comments")%>"><img src="<%=Application("baseroot")&"/common/img/comment_add.png"%>" hspace="0" vspace="0" border="0"></a>
					<%else
						response.Write("<div align='left'>"&langEditor.getTranslated("backend.news.detail.table.label.no_comments")&"</div>")
						end if
						Set objCommento = nothing
						%><br/>
					<%end if%>
				  
					<div id="loading" style="visibility:hidden;display:none;" align="center"><img src="/editor/img/loading.gif" vspace="0" hspace="0" border="0" alt="Loading..." width="200" height="50"></div>
					</form>
				</td>
			</tr>
			</table>
			<br/>
			<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.contenuti.detail.button.inserisci_esci.label")%>" onclick="javascript:sendForm(1);" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.contenuti.detail.button.inserisci.label")%>" onclick="javascript:sendForm(0);" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.commons.back")%>" onclick="javascript:location.href='<%=Application("baseroot")&"/editor/contenuti/ListaNews.asp?cssClass=LN"%>';" />
			<br/><br/>
			
			<%if (Cint(id_news) <> -1) then%>		
				<form action="<%=Application("baseroot") & "/editor/contenuti/DeleteNews.asp"%>" method="post" name="form_cancella_news">
				<input type="hidden" value="<%=id_news%>" name="id_news_to_delete">
				<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.contenuti.detail.button.elimina.label")%>" onclick="javascript:confirmDelete();" />
				</form>
			<%end if
			
			Set objListaTargetPerUser = nothing
			Set objSelNews = Nothing
			%>	
			
		      <form action="<%=Application("baseroot") & "/editor/contenuti/InserisciNews.asp"%>" method="get" name="form_reload_page">
		      <input type="hidden" name="id_news" value="<%=id_news%>">
		      </form>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>