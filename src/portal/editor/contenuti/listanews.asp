<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!-- #include virtual="/editor/include/setListaTargetNews.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script>
function confirmClone(idNews){
	if(confirm('<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.confirm_clone_news")%>')){
		location.href='<%=Application("baseroot") & "/editor/contenuti/clonenews.asp?cssClass=LN&id_news="%>'+idNews;
	}else{
		return;
	}
}
function editContent(idNews){
	location.href='<%=Application("baseroot") & "/editor/contenuti/InserisciNews.asp?cssClass=LN&id_news="%>'+idNews;
}

function deleteContent(id_objref, row,refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.contenuti.detail.js.alert.confirm_del_news")%>?")){		
		ajaxDeleteItem(id_objref,"content",row,refreshrows);
	}
}

function deleteField(id_objref,row,refreshrows){
	if(confirm("<%=langEditor.getTranslated("backend.contenuti.lista.js.alert.delete_field")%>?")){
		ajaxDeleteItem(id_objref,"content_field",row,refreshrows);
	}
}

function showHideDivContentField(element){
	var elementCl = document.getElementById("contenutilist");
	var elementaCl = document.getElementById("acontenutilist");
	var elementCf = document.getElementById("contenutifield");
	var elementaCf = document.getElementById("acontenutifield");

	if(element == 'contenutilist'){
		elementCf.style.visibility = 'hidden';		
		elementCf.style.display = "none";
		elementaCf.className= "";
		elementCl.style.visibility = 'visible';
		elementCl.style.display = "block";
		elementaCl.className= "active";
	}else if(element == 'contenutifield'){
		elementCl.style.visibility = 'hidden';
		elementCl.style.display = "none";
		elementaCl.className= "";
		elementCf.style.visibility = 'visible';		
		elementCf.style.display = "block";
		elementaCf.className= "active";
	}
}

function ajaxViewZoom(idNews, container){
	var dataString;

	if($('#'+container).css("display")=="none"){
		dataString = 'id_news='+ idNews;  
		$.ajax({  
			type: "POST",  
			url: "<%=Application("baseroot") & "/editor/contenuti/ajaxviewcontent.asp"%>",  
			data: dataString,  
			success: function(response) {  
				$('#'+container).html(response); 			
			}
		}); 
	}else{
		$('#'+container).empty();
	}
	$('#'+container).slideToggle();	

	return false; 	
}
</SCRIPT>
</head>
<body onLoad="showHideDivContentField('<%=showTab%>');">
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LN"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
		<table border="0" cellpadding="0" cellspacing="0" align="center" class="filter-table">
		<tr>
		<th align="center"><%=langEditor.getTranslated("backend.contenuti.lista.table.menu.header.txt")%></th>
		</tr>
		<%
		Dim menuFruizioneTmp, iGerLevelTmp, strGerarchiaTmp
		Set menuFruizioneTmp = new MenuClass
		if(isNull(session("strGerTmp")) OR session("strGerTmp") = "" OR (not(isNull(request("resetMenu"))) AND request("resetMenu") = "1")) then 
			session("strGerTmp") = "00" 
			session("contenutiPage") = 1
			numPage = session("contenutiPage")
		end if
		if(request("strGerarchiaTmp") = "") then 
			strGerarchiaTmp = session("strGerTmp") 
		else 
			strGerarchiaTmp = request("strGerarchiaTmp")
			session("strGerTmp") = strGerarchiaTmp
		end if
		iGerLevelTmp = menuFruizioneTmp.getLivello(strGerarchiaTmp)
		
		Set objListCatXNews = CategoriatmpClass.findCategorieByTypeAndMixed(Application("strContentCat"))
		
		for each x in objListCatXNews
			level = menuFruizioneTmp.getLivello(objListCatXNews(x).getCatGerarchia())
			iGerDiff = level - iGerLevelTmp
				
			if(level > 1) then
				iWidth = (level-1) * 10 
				strSubTmpGer=objListCatXNews(x).getCatGerarchia()
				if(level>iGerLevelTmp)then
					numDeltaTmpGer = 0
					if(InStrRev(objListCatXNews(x).getCatGerarchia(),".",-1,1)>0)then
						numDeltaTmpGer = Len(objListCatXNews(x).getCatGerarchia())-(InStrRev(objListCatXNews(x).getCatGerarchia(),".",-1,1)-1)
					end if
					strSubTmpGer = Left(objListCatXNews(x).getCatGerarchia(), Len(objListCatXNews(x).getCatGerarchia())-numDeltaTmpGer)
				end if

				numDeltaSubTmpGer = 0
				if(InStrRev(strSubTmpGer,".",-1,1)>0)then
					numDeltaSubTmpGer = Len(strSubTmpGer)-(InStrRev(strSubTmpGer,".",-1,1)-1)
				end if
				strSubTmpGerFiltered = Left(strSubTmpGer, Len(strSubTmpGer)-numDeltaSubTmpGer)

				if(iGerDiff <= 1) then
					if(iGerDiff<=0)then
					  strSubTmpGer = strSubTmpGerFiltered
					end if
					if (InStr(1, strGerarchiaTmp, strSubTmpGer, 1) > 0) then
						hrefGer = objListCatXNews(x).getCatGerarchia()
						
						'*** checkSelectedCategory
						bolSelectedCat = false
						strSubSelCat = strGerarchiaTmp
						for a=1 to Abs(iGerDiff)
							strSubSelCat = Left(strSubSelCat,InStrRev(strSubSelCat,".",-1,1)-1)
						next

						if(strComp(objListCatXNews(x).getCatGerarchia(), strSubSelCat, 1) = 0) then
							bolSelectedCat = true
						end if%>
						<tr>
							<td><img width="<%=iWidth%>" height="5" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0" align="left"><img src="<%=Application("baseroot")&"/editor/img/folder_explore.png"%>" hspace="0" vspace="0" border="0" align="left"><a href="<%=Application("baseroot") & "/editor/contenuti/ListaNews.asp?page=1&cssClass=LN&target_cat="&objListCatXNews(x).getCatID()&"&strGerarchiaTmp="&hrefGer%>" class="filter-list<%if(bolSelectedCat) then response.Write("-active") end if%>"><%if not(isNull(langEditor.getTranslated(objListCatXNews(x).getCatGerarchia()))) AND not(langEditor.getTranslated(objListCatXNews(x).getCatGerarchia()) = "") then response.write(langEditor.getTranslated(objListCatXNews(x).getCatGerarchia())) else response.Write(objListCatXNews(x).getCatDescrizione()) end if%></a></td>
						</tr>			
					<%end if
				end if
			else				
				numDeltaTmpGer = 0
				if(InStr(1, strGerarchiaTmp, ".", 1) > 0)then
					numDeltaTmpGer = Len(strGerarchiaTmp)-(InStr(1, strGerarchiaTmp, ".", 1)-1)
				end if
				strSubTmpGer = Left(strGerarchiaTmp, Len(strGerarchiaTmp)-numDeltaTmpGer)				
				hrefGer = objListCatXNews(x).getCatGerarchia()%>
				<tr>
					<td><img width="0" height="5" src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" hspace="0" vspace="0" border="0" align="left"><img src="<%=Application("baseroot")&"/editor/img/folder_explore.png"%>" hspace="0" vspace="0" border="0" align="left"><a href="<%=Application("baseroot") & "/editor/contenuti/ListaNews.asp?page=1&cssClass=LN&target_cat="&objListCatXNews(x).getCatID()&"&strGerarchiaTmp="&hrefGer%>" class="filter-list<%if(strComp(objListCatXNews(x).getCatGerarchia(), strSubTmpGer, 1) = 0) then response.Write("-active")%>"><%if not(isNull(langEditor.getTranslated(objListCatXNews(x).getCatGerarchia()))) AND not(langEditor.getTranslated(objListCatXNews(x).getCatGerarchia()) = "") then response.write(langEditor.getTranslated(objListCatXNews(x).getCatGerarchia())) else response.Write(objListCatXNews(x).getCatDescrizione()) end if%></a></td>
				</tr>
			<%end if		
		next	

		Set menuFruizioneTmp = Nothing
		%>
		<tr> 
			<th>&nbsp;</td>
		</tr>
		</table>
		<br>


		<div id="tab-contenuti-field"><a id="acontenutilist" <%if(showtab="contenutilist")then response.write("class=active") end if%> href="javascript:showHideDivContentField('contenutilist');"><%=langEditor.getTranslated("backend.contenuti.lista.table.header.label_contenuti_list")%></a><a id="acontenutifield" <%if(showtab="contenutifield")then response.write("class=active") end if%> href="javascript:showHideDivContentField('contenutifield');"><%=langEditor.getTranslated("backend.contenuti.lista.table.header.label_contenuti_field")%></a></div>		
		<div id="ajaxresp" align="center" style="background-color:#FFFF00; border:1px solid #000000; color:#000000; display:none;"></div>
		<div id="contenutilist" style="visibility:visible;display:block;margin:0px;padding:0px;">
		<table border="0" cellpadding="0" cellspacing="0" class="principal">
		      <tr> 
			<th colspan="4">&nbsp;</td>
			<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.title")%>&nbsp;<a href="<%=Application("baseroot") & "/editor/contenuti/ListaNews.asp?order_by=1&page="&numPage&"&strGerarchiaTmp="&request("strGerarchiaTmp")&"&target_cat="&request("target_cat")&"&items="&itemsXpage%>"><img src="<%=Application("baseroot")&"/editor/img/order_top.gif"%>" alt="<%=langEditor.getTranslated("backend.commons.alt.order_by_asc")%>" hspace="2" vspace="0" border="0"></a><a href="<%=Application("baseroot") & "/editor/contenuti/ListaNews.asp?order_by=2&page="&numPage&"&strGerarchiaTmp="&request("strGerarchiaTmp")&"&target_cat="&request("target_cat")&"&items="&itemsXpage%>"><img src="<%=Application("baseroot")&"/editor/img/order_bottom.gif"%>" alt="<%=langEditor.getTranslated("backend.commons.alt.order_by_desc")%>" hspace="2" vspace="0" border="0"></a></th>
			<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.pub_date")%>&nbsp;<a href="<%=Application("baseroot") & "/editor/contenuti/ListaNews.asp?order_by=11&page="&numPage&"&strGerarchiaTmp="&request("strGerarchiaTmp")&"&target_cat="&request("target_cat")&"&items="&itemsXpage%>"><img src="<%=Application("baseroot")&"/editor/img/order_top.gif"%>" alt="<%=langEditor.getTranslated("backend.commons.alt.order_by_asc")%>" hspace="2" vspace="0" border="0"></a><a href="<%=Application("baseroot") & "/editor/contenuti/ListaNews.asp?order_by=12&page="&numPage&"&strGerarchiaTmp="&request("strGerarchiaTmp")&"&target_cat="&request("target_cat")&"&items="&itemsXpage%>"><img src="<%=Application("baseroot")&"/editor/img/order_bottom.gif"%>" alt="<%=langEditor.getTranslated("backend.commons.alt.order_by_desc")%>" hspace="2" vspace="0" border="0"></a></th>
			<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.stato")%></th>
			<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.category")%></th>
			<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.lang")%></th>
		      </tr>
			  
				<%
				Dim hasNews
				hasNews = false
				on error Resume Next
					if(Cint(strRuoloLogged) = Application("admin_role")) then
					'response.write("start: "&Time()&"<br>")
						Set objListaNews = objNews.findNewsSlim(null, null, null, null, objListaTargetCatTmp, objListaTargetLangTmp, null, null, null, order_news_by, false, false)
					'response.write("end: "&Time()&"<br>")		
					else
						Set objListaNews = objNews.findNewsSlim(null, objUserLogged.getUserID(), null, null, objListaTargetCatTmp, objListaTargetLangTmp, null, null, null, order_news_by, false, false)		
					end if
					
					if(objListaNews.Count > 0) then
						hasNews = true
					end if
					
				if Err.number <> 0 then
					'response.write(Err.description)
				end if	
				
				if(hasNews) then
				
					Dim intCount
					intCount = 0
					
					Dim newsCounter, iIndex, objTmpNews, objTmpNewsKey, FromNews, ToNews, Diff, objTarget
					iIndex = objListaNews.Count
					FromNews = ((numPageNews * itemsXpageNews) - itemsXpageNews)
					Diff = (iIndex - ((numPageNews * itemsXpageNews)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToNews = iIndex - Diff
					
					totPages = iIndex\itemsXpageNews
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpageNews <> 0) AND not ((totPages * itemsXpageNews) >= iIndex)) then
						totPages = totPages +1	
					end if		
							
					objTmpNews = objListaNews.Items
					objTmpNewsKey=objListaNews.Keys
					objTarget = null
					
					Dim styleRow, styleRow2
					styleRow2 = "table-list-on"
					
							
					for newsCounter = FromNews to ToNews
						styleRow = "table-list-off"
						if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if
						Set objFilteredNews = objTmpNews(newsCounter)
						objFilteredNews.setListaTarget(objFilteredNews.getTargetPerNews(objFilteredNews.getNewsID()))
						Set objTarget = objFilteredNews.getListaTarget()
						%>		
						<tr class="<%=styleRow%>" id="tr_delete_list_<%=intCount%>">
						<td align="center" width="25"><a href="javascript:confirmClone('<%=objFilteredNews.getNewsID()%>');"><img src="<%=Application("baseroot")&"/editor/img/page_white_copy.png"%>" alt="<%=langEditor.getTranslated("backend.contenuti.lista.table.alt.clone")%>" hspace="2" vspace="0" border="0"></a></td>
						<td align="center" width="25"><!--<a href="<%'=Application("baseroot") & "/editor/contenuti/VisualizzaNews.asp?cssClass=LN&id_news=" & objFilteredNews.getNewsID()%>">--><img style="cursor:pointer;" id="view_zoom_<%=intCount%>" src="<%=Application("baseroot")&"/editor/img/zoom.png"%>" alt="<%=langEditor.getTranslated("backend.contenuti.lista.table.alt.view")%>" hspace="2" vspace="0" border="0"><!--</a>--></td>
						<td align="center" width="25"><a href="javascript:editContent('<%=objFilteredNews.getNewsID()%>');"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.contenuti.lista.table.alt.modify")%>" hspace="2" vspace="0" border="0"></a></td>
						<td align="center" width="25"><a href="javascript:deleteContent(<%=objFilteredNews.getNewsID()%>,'tr_delete_list_<%=intCount%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.contenuti.detail.button.elimina.label")%>" hspace="2" vspace="0" border="0"></a></td>
						<td nowrap width="30%">						
						<strong><div class="ajax" id="view_title_<%=intCount%>" onmouseover="javascript:showHide('view_title_<%=intCount%>','edit_title_<%=intCount%>','title_<%=intCount%>',500, false);"><%=objFilteredNews.getTitolo()%></div></strong>
						<div class="ajax" id="edit_title_<%=intCount%>"><textarea class="formfieldAjaxArea" id="title_<%=intCount%>" name="titolo" onmouseout="javascript:restoreField('edit_title_<%=intCount%>','view_title_<%=intCount%>','title_<%=intCount%>','content',<%=objFilteredNews.getNewsID()%>,1,<%=intCount%>);"><%=objFilteredNews.getTitolo()%></textarea></div>
						<script>
						$("#edit_title_<%=intCount%>").hide();
						</script>
						</td>
						<td width="17%">
						<div class="ajax" id="view_news_data_pub_<%=intCount%>" onmouseover="javascript:showHide('view_news_data_pub_<%=intCount%>','edit_news_data_pub_<%=intCount%>','news_data_pub_<%=intCount%>',500, true);"><%=FormatDateTime(objFilteredNews.getDataPubNews(),2)&" "&FormatDateTime(objFilteredNews.getDataPubNews(),vbshorttime)%></div>
						<div class="ajax" id="edit_news_data_pub_<%=intCount%>"><input type="text" class="formfieldAjax" id="news_data_pub_<%=intCount%>" name="news_data_pub" onchange="javascript:updateField('edit_news_data_pub_<%=intCount%>','view_news_data_pub_<%=intCount%>','news_data_pub_<%=intCount%>','content',<%=objFilteredNews.getNewsID()%>,1,<%=intCount%>);" value="<%=FormatDateTime(objFilteredNews.getDataPubNews(),2)&" "&FormatDateTime(objFilteredNews.getDataPubNews(),vbshorttime)%>"></div>
						<script>
						
						$(function() {
							$('#news_data_pub_<%=intCount%>').datetimepicker({
								showButtonPanel: false,
								dateFormat: 'dd/mm/yy',
								timeFormat: 'hh.mm'
							});
							$('#ui-datepicker-div').hide();							
						});

						$("#edit_news_data_pub_<%=intCount%>").hide();
						</script>
						</td>
						<td width="10%">
						<div class="ajax" id="view_stato_news_<%=intCount%>" onmouseover="javascript:showHide('view_stato_news_<%=intCount%>','edit_stato_news_<%=intCount%>','stato_news_<%=intCount%>',500, true);">
						<%
						Select Case objFilteredNews.getStato()
						Case 0
							response.write(langEditor.getTranslated("backend.contenuti.lista.table.select.option.edit"))
						Case 1
							response.write(langEditor.getTranslated("backend.contenuti.lista.table.select.option.public"))
						Case Else
						End Select%>
						</div>
						<div class="ajax" id="edit_stato_news_<%=intCount%>">
						<select name="stato_news" class="formfieldAjaxSelect" id="stato_news_<%=intCount%>" onblur="javascript:updateField('edit_stato_news_<%=intCount%>','view_stato_news_<%=intCount%>','stato_news_<%=intCount%>','content',<%=objFilteredNews.getNewsID()%>,2,<%=intCount%>);">
						<option value="0"<%if (0=Cint(objFilteredNews.getStato())) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.contenuti.lista.table.select.option.edit")%></option>	
						<option value="1"<%if (1=Cint(objFilteredNews.getStato())) then response.Write(" selected")%>><%=langEditor.getTranslated("backend.contenuti.lista.table.select.option.public")%></option>	
						</select>	
						</div>
						<script>
						$("#edit_stato_news_<%=intCount%>").hide();
						</script>
						</td>
						<td nowrap>
						<%	
						Dim objCategorieXNews	
						if (Instr(1, typename(objTarget), "dictionary", 1) > 0) then
							for each y in objTarget.Keys
								if (objTarget(y).getTargetType() = 1) then
									Set objCategorieXNews = CategoriatmpClass.findCategorieByTargetID(y)
									if not (isNull(objCategorieXNews)) then
										for each z in objCategorieXNews.Keys
											response.write ("<a class=""link-change-cat"" href="""&Application("baseroot") & "/editor/contenuti/ListaNews.asp?page=1&cssClass=LN&target_cat="&objCategorieXNews(z).getCatID()&"&strGerarchiaTmp="&objCategorieXNews(z).getCatGerarchia()&""" title="""&langEditor.getTranslated("backend.contenuti.lista.table.alt.filter_cat")&""">" & objCategorieXNews(z).getCatDescrizione() & "</a><br>")
										next
									end if
									Set objCategorieXNews = nothing
								end if									
							next	
						end if%>
						</td>
						<td nowrap>
						<%		
						if (Instr(1, typename(objTarget), "dictionary", 1) > 0) then
							tcount = 1
							for each y in objTarget.Keys
								if (objTarget(y).getTargetType() = 3) then
									imtTitle = Replace(objTarget(y).getTargetDescrizione(), "lang_", "", 1, -1, 1)
									if not(langEditor.getTranslated("portal.header.label.desc_lang."&Replace(objTarget(y).getTargetDescrizione(), "lang_", "", 1, -1, 1)) = "") then imtTitle = langEditor.getTranslated("portal.header.label.desc_lang."&Replace(objTarget(y).getTargetDescrizione(), "lang_", "", 1, -1, 1)) end if%>
									<img width="16" height="11" border="0" style="padding-right:0px;" alt="<%=imtTitle%>" title="<%=imtTitle%>" src="/editor/img/flag/flag-<%=Replace(objTarget(y).getTargetDescrizione(), "lang_", "", 1, -1, 1)%>.png"><%if(tcount MOD 4 =0)then response.write("<br/>") end if%>
									<%tcount = tcount+1
								end if		
							next		
							Set objTarget = nothing
						end if%>
						</td>
						</tr>

						<tr class="preview_row">
						<td colspan="9">
						<div id="view_content_<%=intCount%>"></div>
						<script>
						$("#view_content_<%=intCount%>").hide();
						$('#view_zoom_<%=intCount%>').click(function(){ajaxViewZoom('<%=objFilteredNews.getNewsID()%>', 'view_content_<%=intCount%>');});
						</script>	
						</td>
						</tr>

						<%intCount = intCount +1
						Set objFilteredNews = nothing
					next
					Set objListaNews = nothing
					%>
				  
				  <tr> 
					<form action="<%=Application("baseroot") & "/editor/contenuti/ListaNews.asp"%>" method="post" name="item_x_page">
					<th colspan="9" align="left">
					<input type="hidden" value="<%=order_news_by%>" name="order_by">
					<input type="hidden" value="<%=target_cat_param%>" name="target_cat">
					<input type="hidden" value="<%=request("strGerarchiaTmp")%>" name="strGerarchiaTmp">				
					<input type="text" name="itemsNews" class="formFieldTXTNumXPage" value="<%=itemsXpageNews%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
					<%		
					'**************** richiamo paginazione
					call PaginazioneFrontend(totPages, numPageNews, strGerarchia, "/editor/contenuti/ListaNews.asp", "&order_by="&order_news_by&"&target_cat="&target_cat_param&"&itemsNews="&itemsXpageNews&"&strGerarchiaTmp="&request("strGerarchiaTmp"))%>				
					</th>
					</form>
              		</tr>
              	<%end if
				
			Set objListCatXNews = Nothing
			Set CategoriatmpClass = Nothing
			Set objNews = Nothing%>
		</table>
		<br/>
		<div style="float:left;">
		<form action="<%=Application("baseroot") & "/editor/contenuti/InserisciNews.asp?cssClass=LN"%>" method="post" name="form_crea">
			<input type="hidden" value="-1" name="id_news">
			<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.contenuti.lista.button.inserisci.label")%>" onclick="javascript:document.form_crea.submit();" />
		</form>
		</div>
		<div style="float:left;">
		&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.contenuti.lista.button.label.download_excel")%>" onclick="javascript:openWinExcel('<%=Application("baseroot")&"/editor/report/CreateDownFileExcel.asp"%>','crea_excel',400,400,100,100);" />
		</div>
		<div>
		&nbsp;<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("portal.templates.commons.label.see_comments_news")%>" onclick="javascript:openWin('<%=Application("baseroot")&"/editor/include/popupCommentManager.asp?element_type=1"%>','popupallegati',400,400,100,100);" />
		</div>		
		</div>
		<div id="contenutifield" style="visibility:hidden;margin:0px;padding:0px;">
			<table border="0" cellpadding="0" cellspacing="0" class="principal" align="top">
			<tr> 
				<th colspan="2">&nbsp;</th>
				<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.description")%></th>
				<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.group")%></th>
				<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.order")%></th>
				<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.type")%></th>
				<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.type_content")%></th>
				<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.required")%></th>
				<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.enabled")%></th>
				<th><%=langEditor.getTranslated("backend.contenuti.lista.table.header.editable")%></th>
			</tr>
				<%
				Dim bolHasObj
				bolHasObj = false
				intCount = 0
				iIndex = 0				

				On Error Resume Next
				Set objListaField = objContentField.getListContentField(null)
				if(objListaField.Count > 0) then		
					bolHasObj = true
				end if

				if Err.number <> 0 then
					bolHasObj = false
				end if			
				
				if(bolHasObj) then
					Dim tmpObjField				
					Dim objTmpField, objTmpFieldKey, FromField, ToField
					iIndex = objListaField.Count
					FromField = ((numPageField * itemsXpageField) - itemsXpageField)
					Diff = (iIndex - ((numPageField * itemsXpageField)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToField = iIndex - Diff
					
					totPages = iIndex\itemsXpageField
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpageField <> 0) AND not ((totPages * itemsXpageField) >= iIndex)) then
						totPages = totPages +1	
					end if		
					
					styleRow2 = "table-list-on"
					
					objTmpField = objListaField.Items
					objTmpFieldKey=objListaField.Keys		
					for newsCounter = FromField to ToField
						styleRow = "table-list-off"
						if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>
						<form action="<%=Application("baseroot") & "/editor/contenuti/InserisciField.asp"%>" method="post" name="form_lista_field_<%=intCount%>">
						<input type="hidden" value="<%=objTmpFieldKey(newsCounter)%>" name="id_field">
						<input type="hidden" value="" name="delete_field"> 
						<input type="hidden" value="LN" name="cssClass">	
						</form>			
						<tr class="<%=styleRow%>" id="tr_delete_field_list_<%=intCount%>">				
							<%Set tmpObjField = objTmpField(newsCounter)%>
							<td align="center" width="25"><a href="javascript:document.form_lista_field_<%=intCount%>.submit();"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=langEditor.getTranslated("backend.contenuti.lista.table.alt.modify_field")%>" hspace="2" vspace="0" border="0"></a></td>
							<td align="center" width="25"><a href="javascript:deleteField(<%=objTmpFieldKey(newsCounter)%>,'tr_delete_field_list_<%=intCount%>','tr_delete_field_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=langEditor.getTranslated("backend.contenuti.lista.table.alt.delete_field")%>" hspace="5" vspace="0" border="0"></a></td>										
							<td width="16%">						
							<div class="ajax" id="view_description_<%=intCount%>" onmouseover="javascript:showHide('view_description_<%=intCount%>','edit_description_<%=intCount%>','description_<%=intCount%>',500, false);"><%=tmpObjField.getDescription()%></div>
							<div class="ajax" id="edit_description_<%=intCount%>"><input type="text" class="formfieldAjax" id="description_<%=intCount%>" name="description" onmouseout="javascript:restoreField('edit_description_<%=intCount%>','view_description_<%=intCount%>','description_<%=intCount%>','content_field',<%=tmpObjField.getID()%>,1,<%=intCount%>);" value="<%=tmpObjField.getDescription()%>" onkeypress="javascript:return notSpecialChar(event);"></div>
							<script>
							$("#edit_description_<%=intCount%>").hide();
							</script>
							</td>
							<td width="18%">						
							<div class="ajax" id="view_id_group_<%=intCount%>" onmouseover="javascript:showHide('view_id_group_<%=intCount%>','edit_id_group_<%=intCount%>','id_group_<%=intCount%>',500, true);"><%=tmpObjField.getObjGroup().getDescription()%></div>
							<div class="ajax" id="edit_id_group_<%=intCount%>">
							<select name="id_group" class="formfieldAjaxSelect" id="id_group_<%=intCount%>" onblur="javascript:updateField('edit_id_group_<%=intCount%>','view_id_group_<%=intCount%>','id_group_<%=intCount%>','content_field',<%=tmpObjField.getID()%>,2,<%=intCount%>);">
							<%
							On Error resume next
							Set objFieldGroup = New ContentFieldGroupClass
							Dim objDispFGroup
							Set objDispFGroup = objFieldGroup.getListContentFieldGroup()
							Set objFieldGroup = nothing

							if (Instr(1, typename(objDispFGroup), "dictionary", 1) > 0) then
							for each x in objDispFGroup%>
							<option value="<%=x%>" <%if (tmpObjField.getIdGroup() = x) then response.Write("selected")%>><%if not(langEditor.getTranslated("backend.contenuti.detail.table.label.group."&objDispFGroup(x).getDescription()) = "") then response.write(langEditor.getTranslated("backend.contenuti.detail.table.label.group."&objDispFGroup(x).getDescription())) else response.write(objDispFGroup(x).getDescription()) end if%></option>
							<%next
							end if
							Set objDispFGroup = nothing
							if(Err.number <>0)then
							'response.write(Err.description)
							end if%>
							</select>
							</div>
							<script>
							$("#edit_id_group_<%=intCount%>").hide();
							</script>
							</td>
							<td>	
							<div class="ajax" id="view_order_<%=intCount%>" onmouseover="javascript:showHide('view_order_<%=intCount%>','edit_order_<%=intCount%>','order_<%=intCount%>',500, false);"><%=tmpObjField.getOrder()%></div>
							<div class="ajax" id="edit_order_<%=intCount%>"><input type="text" class="formfieldAjaxShort" id="order_<%=intCount%>" name="order" onmouseout="javascript:restoreField('edit_order_<%=intCount%>','view_order_<%=intCount%>','order_<%=intCount%>','content_field',<%=tmpObjField.getID()%>,1,<%=intCount%>);" value="<%=tmpObjField.getOrder()%>" maxlength="3" onkeypress="javascript:return isInteger(event);"></div>
							<script>
							$("#edit_order_<%=intCount%>").hide();
							</script>
							</td>
							<td><%=objContentField.findTypeFieldById(tmpObjField.getTypeField())%></td>
							<td><%=objContentField.findTypeContentById(tmpObjField.getTypeContent())%></td>
							<td>
							<%
							if (strComp("1", tmpObjField.getRequired(), 1) = 0) then 
								response.Write(langEditor.getTranslated("backend.commons.yes"))
							else 
								response.Write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</td>
							<td>
							<%
							if (strComp("1", tmpObjField.getEnabled(), 1) = 0) then 
								response.Write(langEditor.getTranslated("backend.commons.yes"))
							else 
								response.Write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</td>	
							<td>
							<%
							if (strComp("1", tmpObjField.getEditable(), 1) = 0) then 
								response.Write(langEditor.getTranslated("backend.commons.yes"))
							else 
								response.Write(langEditor.getTranslated("backend.commons.no"))
							end if
							%>
							</td>		
						</tr>			
						<%intCount = intCount +1
					next
					Set tmpObjField = nothing
					Set objListaField = nothing%>
		      <tr> 
			<form action="<%=Application("baseroot") & "/editor/contenuti/Listanews.asp"%>" method="post" name="item_x_page_field">
			<input type="hidden" value="contenutifield" name="showtab">
			<th colspan="10">
					<input type="text" name="itemsField" class="formFieldTXTNumXPage" value="<%=itemsXpageField%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
					<%						
					'**************** richiamo paginazione
					call PaginazioneFrontend(totPages, numPageField, strGerarchia, "/editor/contenuti/Listanews.asp", "&itemsField="&itemsXpageField&"&showtab=contenutifield")
					%>
			</th>
			</form>
		      </tr>
		      <%end if%>
		    </table>
			<br/>
			<form action="<%=Application("baseroot") & "/editor/contenuti/InserisciField.asp"%>" method="post" name="form_crea_field">
			<input type="hidden" value="-1" name="id_field">
			<input type="hidden" value="LN" name="cssClass">
			<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.contenuti.lista.button.inserisci_field.label")%>" onclick="javascript:document.form_crea_field.submit();" />
			</form>			
		</div>
		<%
		Set objContentField = nothing
		%>
		<br/><br/>		
	</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>