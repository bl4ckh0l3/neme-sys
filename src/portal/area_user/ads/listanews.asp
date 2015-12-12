<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->
<!-- #include file="include/init.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("backend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/area_user.css"%>" type="text/css">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<!-- gestione degli update fields via ajax -->
<script language="JavaScript">
function sendAjaxCommand(field_name, field_val, objtype, id_objref, listCounter, field){
	var query_string = "field_name="+field_name+"&field_val="+encodeURIComponent(field_val)+"&objtype="+objtype+"&id_objref="+id_objref;
	//alert("query_string: "+query_string);
	var resp = false;

	$.ajax({
		async: false,
		type: "GET",
		cache: false,
		url: "<%=Application("baseroot") & "/area_user/ads/ajaxupdate.asp"%>",
		data: query_string,
		success: function(response) {
			resp = true;

			// il codice seguente server per inviare il contatore dell'oggetto modificato nella lista
			// per chiamare la funzione specifica di ogni pagina, per modificare elementi della pagina accessori
			if(typeof changeRowListData == 'function'){				
				changeRowListData(listCounter, objtype, field);
			}
		},
		error: function() {
			$("#ajaxresp").empty();
			$("#ajaxresp").append("<%=lang.getTranslated("backend.commons.fail_updated_field")%>");
			$("#ajaxresp").fadeIn(1500,"linear");
			$("#ajaxresp").fadeOut(600,"linear");
			resp = false;
		}
	});

	return resp;
}


function ajaxDeleteItem(id_objref,objtype,row,refreshrows){
	var query_string = "id_objref="+id_objref+"&objtype="+objtype;
	
	$.ajax({
		async: false,
		type: "GET",
		cache: false,
		url: "<%=Application("baseroot") & "/area_user/ads/ajaxdelete.asp"%>",
		data: query_string,
		success: function(response) {
			if(response ==""){
				var classon = "table-list-on";
				var classoff = "table-list-off";
				var counter = 1;
        
				$('#'+row).remove();	
				
				$("tr[id*='"+refreshrows+"']").each(function(){
					if(counter % 2 == 0){
						$(this).attr("class",classoff);
					}else{
						$(this).attr("class",classon);
					}
					counter+=1;
				});	
				
			}else{
				location.href='<%=Application("baseroot")&Application("error_page")&"?error="%>'+response;				
			}
		},
		error: function() {
			$("#ajaxresp").empty();
			$("#ajaxresp").append("<%=lang.getTranslated("backend.commons.fail_delete_item")%>");
			$("#ajaxresp").fadeIn(1500,"linear");
			$("#ajaxresp").fadeOut(600,"linear");
		}
	});
}


var field_lock = false;
var has_focus = false;
var orig_val;
function showHide(fieldHide, fieldShow, field, mode, focus){
	var timer = 1500;
	if(!field_lock){
		$("#"+fieldHide).hide();
		$("#"+fieldShow).show();
		//$("#"+fieldShow).show(mode);
		if(focus){
			$('#'+field).focus();
			timer = 2000;
		}
		orig_val = $('#'+field).val();
		field_lock = true;

		setTimeout(function(){resetFieldFocus(fieldShow, fieldHide, field, orig_val, focus);}, timer);
	}
}

function updateField(fieldHide, fieldShow, field, objtype, id_objref, field_type, listCounter){
	var edit_val_ch = $('#'+field).val();
	var field_name = $('#'+field).attr("name");
	var resp = false;
  
  //alert("updateField - edit_val_ch: "+edit_val_ch);
  //alert("updateField - field_name: "+field_name);

	if(edit_val_ch != orig_val){
		resp = sendAjaxCommand(field_name, edit_val_ch, objtype, id_objref, listCounter, field);
	}else{
		orig_val = "";
	}	

	if(resp){
		$("#"+fieldShow).empty();
		if(field_type==2){
			$("#"+fieldShow).append($('#'+field+' :selected').text());		
		}else{
			$("#"+fieldShow).append(edit_val_ch);			
		}
	}

	$("#"+fieldHide).hide();
	$("#"+fieldShow).show();
	field_lock = false;
	has_focus = false;
}

function restoreField(fieldHide, fieldShow, field, objtype, id_objref, field_type, listCounter){
	var edit_val_ch = $('#'+field).val();
  
  //alert("restoreField - edit_val_ch: "+edit_val_ch);
	
	if(edit_val_ch != orig_val){
		updateField(fieldHide, fieldShow, field, objtype, id_objref, field_type, listCounter)
	}

	$("#"+fieldHide).hide();
	$("#"+fieldShow).show();
	field_lock = false;
	has_focus = false;
}

function resetFieldFocus(fieldHide, fieldShow, field, orig_val, focus){
	if(orig_val==$('#'+field).val()){
		if(has_focus==false){	
			if(focus){
				$("#"+field).blur();
				has_focus = false;
			}else{
				$("#"+fieldHide).hide();
				$("#"+fieldShow).show();
				field_lock = false;
				has_focus = false;
			}	
		}
	}
}

function setFocusField(){
	has_focus=true;
}

$(document).ready(function() {
	$("input[type='text']").click( function() {setFocusField();});
	$("textarea").click(function() {setFocusField();});
	$("select").click(function() {setFocusField();});
});
</script>
<script>
/*<!--nsys-usr-lnews1-->*/
function createAds(idNews){
	location.href='<%=Application("baseroot") & "/area_user/ads/InserisciAds.asp?id_news="%>'+idNews;
}
/*<!---nsys-usr-lnews1-->*/

function editContent(idNews){
	location.href='<%=Application("baseroot") & "/area_user/ads/InserisciNews.asp?cssClass=LN&id_news="%>'+idNews;
}

function deleteContent(id_objref, row,refreshrows){
	if(confirm("<%=lang.getTranslated("backend.contenuti.detail.js.alert.confirm_del_news")%>?")){		
		ajaxDeleteItem(id_objref,"content",row,refreshrows);
	}
}
</SCRIPT>
</head>
<body>
<!-- #include virtual="/public/layout/area_user/grid_top.asp" -->

		<div id="ajaxresp" align="center" style="background-color:#FFFF00; border:1px solid #000000; color:#000000; display:none;"></div>
		<table border="0" cellpadding="0" cellspacing="0" class="principal">
		      <tr> 
			<!--nsys-usr-lnews2--><th colspan="3">&nbsp;</td><!---nsys-usr-lnews2-->
			<th><%=lang.getTranslated("backend.contenuti.lista.table.header.title")%>&nbsp;<a href="<%=Application("baseroot") & "/area_user/ads/ListaNews.asp?order_by=1&page="&numPage&"&strGerarchiaTmp="&request("strGerarchiaTmp")&"&target_cat="&request("target_cat")&"&items="&itemsXpage%>"><img src="<%=Application("baseroot")&"/editor/img/order_top.gif"%>" alt="<%=lang.getTranslated("backend.commons.alt.order_by_asc")%>" hspace="2" vspace="0" border="0"></a><a href="<%=Application("baseroot") & "/area_user/ads/ListaNews.asp?order_by=2&page="&numPage&"&strGerarchiaTmp="&request("strGerarchiaTmp")&"&target_cat="&request("target_cat")&"&items="&itemsXpage%>"><img src="<%=Application("baseroot")&"/editor/img/order_bottom.gif"%>" alt="<%=lang.getTranslated("backend.commons.alt.order_by_desc")%>" hspace="2" vspace="0" border="0"></a></th>
			<th><%=lang.getTranslated("backend.contenuti.lista.table.header.pub_date")%>&nbsp;<a href="<%=Application("baseroot") & "/area_user/ads/ListaNews.asp?order_by=11&page="&numPage&"&strGerarchiaTmp="&request("strGerarchiaTmp")&"&target_cat="&request("target_cat")&"&items="&itemsXpage%>"><img src="<%=Application("baseroot")&"/editor/img/order_top.gif"%>" alt="<%=lang.getTranslated("backend.commons.alt.order_by_asc")%>" hspace="2" vspace="0" border="0"></a><a href="<%=Application("baseroot") & "/area_user/ads/ListaNews.asp?order_by=12&page="&numPage&"&strGerarchiaTmp="&request("strGerarchiaTmp")&"&target_cat="&request("target_cat")&"&items="&itemsXpage%>"><img src="<%=Application("baseroot")&"/editor/img/order_bottom.gif"%>" alt="<%=lang.getTranslated("backend.commons.alt.order_by_desc")%>" hspace="2" vspace="0" border="0"></a></th>
			<th><%=lang.getTranslated("backend.contenuti.lista.table.header.stato")%></th>
			<th><%=lang.getTranslated("backend.contenuti.lista.table.header.category")%></th>
			<th><%=lang.getTranslated("backend.contenuti.lista.table.header.lang")%></th>
		      </tr>
			  
				<%
				Dim hasNews
				hasNews = false
				on error Resume Next
				Set objListaNews = objNews.findNewsSlim(null, Session("objUtenteLogged"), null, null, objListaTargetCatTmp, objListaTargetLangTmp, null, null, null, order_news_by, false, false)
				
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
					FromNews = ((numPage * itemsXpage) - itemsXpage)
					Diff = (iIndex - ((numPage * itemsXpage)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToNews = iIndex - Diff
					
					totPages = iIndex\itemsXpage
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
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
						<!--nsys-usr-lnews3--><td align="center" width="25"><%if(Application("enable_ads")=1)then%><a href="javascript:createAds('<%=objFilteredNews.getNewsID()%>');"><img src="<%=Application("baseroot")&"/editor/img/vcard_add.png"%>" alt="<%=lang.getTranslated("backend.contenuti.lista.table.alt.ads")%>" title="<%=lang.getTranslated("backend.contenuti.lista.table.alt.ads")%>" hspace="2" vspace="0" border="0"></a><%end if%></td><!---nsys-usr-lnews3-->
						<td align="center" width="25"><a href="javascript:editContent('<%=objFilteredNews.getNewsID()%>');"><img src="<%=Application("baseroot")&"/editor/img/pencil.png"%>" alt="<%=lang.getTranslated("backend.contenuti.lista.table.alt.modify")%>" title="<%=lang.getTranslated("backend.contenuti.lista.table.alt.modify")%>" hspace="2" vspace="0" border="0"></a></td>
						<td align="center" width="25"><a href="javascript:deleteContent(<%=objFilteredNews.getNewsID()%>,'tr_delete_list_<%=intCount%>','tr_delete_list_');"><img src="<%=Application("baseroot")&"/editor/img/cancel.png"%>" alt="<%=lang.getTranslated("backend.contenuti.detail.button.elimina.label")%>" title="<%=lang.getTranslated("backend.contenuti.detail.button.elimina.label")%>" hspace="2" vspace="0" border="0"></a></td>
						<td nowrap width="27%"><strong><%=objFilteredNews.getTitolo()%></strong></td>
						<td width="15%"><%=FormatDateTime(objFilteredNews.getDataPubNews(),2)&" "&FormatDateTime(objFilteredNews.getDataPubNews(),vbshorttime)%></td>
						<td width="15%">
						<div class="ajax" id="view_stato_news_<%=intCount%>" onmouseover="javascript:showHide('view_stato_news_<%=intCount%>','edit_stato_news_<%=intCount%>','stato_news_<%=intCount%>',500, true);">
						<%
						Select Case objFilteredNews.getStato()
						Case 0
							response.write(lang.getTranslated("backend.contenuti.lista.table.select.option.edit"))
						Case 1
							response.write(lang.getTranslated("backend.contenuti.lista.table.select.option.public"))
						Case Else
						End Select%>
						</div>
						<div class="ajax" id="edit_stato_news_<%=intCount%>">
						<select name="stato_news" class="formfieldAjaxSelect" id="stato_news_<%=intCount%>" onblur="javascript:updateField('edit_stato_news_<%=intCount%>','view_stato_news_<%=intCount%>','stato_news_<%=intCount%>','content',<%=objFilteredNews.getNewsID()%>,2,<%=intCount%>);">
						<option value="0"<%if (0=Cint(objFilteredNews.getStato())) then response.Write(" selected")%>><%=lang.getTranslated("backend.contenuti.lista.table.select.option.edit")%></option>	
						<option value="1"<%if (1=Cint(objFilteredNews.getStato())) then response.Write(" selected")%>><%=lang.getTranslated("backend.contenuti.lista.table.select.option.public")%></option>	
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
											response.write (objCategorieXNews(z).getCatDescrizione() & "<br>")
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
						<%intCount = intCount +1
						Set objFilteredNews = nothing
					next
					Set objListaNews = nothing
					%>
				  
				  <tr> 
					<!--nsys-usr-lnews4--><th colspan="8" align="left"><!---nsys-usr-lnews4-->
					<div style="float:left;padding-right:4px;">
					<form action="<%=Application("baseroot") & "/area_user/ads/ListaNews.asp"%>" method="post" name="item_x_page">
					<input type="hidden" value="<%=order_news_by%>" name="order_by">
					<input type="hidden" value="<%=target_cat_param%>" name="target_cat">
					<input type="hidden" value="<%=request("strGerarchiaTmp")%>" name="strGerarchiaTmp">				
					<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=lang.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
					</form>
					</div>
					<div>
					<%		
					'**************** richiamo paginazione
					call PaginazioneFrontend(totPages, numPage, strGerarchia, "/area_user/ads/ListaNews.asp", "order_by="&order_news_by&"&target_cat="&target_cat_param&"&items="&itemsXpage&"&strGerarchiaTmp="&request("strGerarchiaTmp"))%>
					</div>					
					</th>
              		</tr>
              	<%end if
				
			Set objListCatXNews = Nothing
			Set CategoriatmpClass = Nothing
			Set objNews = Nothing%>
		</table>
		<br/>
		<div style="float:left;">
		<form action="<%=Application("baseroot") & "/area_user/ads/InserisciNews.asp"%>" method="post" name="form_crea">
			<input type="hidden" value="-1" name="id_news">
			<input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=lang.getTranslated("backend.contenuti.lista.button.inserisci.label")%>" onclick="javascript:document.form_crea.submit();" />
		</form>
		</div>

<!-- #include virtual="/public/layout/area_user/grid_bottom.asp" -->
</body>
</html>