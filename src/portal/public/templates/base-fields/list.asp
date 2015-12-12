<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->
<!-- #include file="include/init1.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=pageTemplateTitle%></title>
<META name="description" CONTENT="<%=metaDescription%>">
<META name="keywords" CONTENT="<%=metaKeyword%>">
<META name="autore" CONTENT="Neme-sys; email:info@neme-sys.org">
<META http-equiv="Content-Type" CONTENT="text/html; charset=utf-8">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<%if not(isNull(strCSS)) ANd not(strCSS = "") then%><link rel="stylesheet" href="<%=Application("baseroot") & strCSS%>" type="text/css"><%end if%>
<script language="Javascript">  
function openDetailContentPage(strAction, strGerarchia, numIdNews, numPageNum){
    document.form_detail_link_news.action=strAction;
    document.form_detail_link_news.gerarchia.value=strGerarchia;
    document.form_detail_link_news.id_news.value=numIdNews;
    document.form_detail_link_news.modelPageNum.value=numPageNum;
    document.form_detail_link_news.submit();
}

function resetFilters(){
	$("select[name*='field_']").each( function(){
		$(this).empty();
	});
	$('#fields_filter').val("0");
	$('#reset_fields_filter').val("1");
	form_field_filter.submit();
}
</script>
</head>
<body>
<div id="warp">
	<!-- #include virtual="/public/layout/include/header.inc" -->	
	<div id="container">	
		<!-- include virtual="/public/layout/include/menu_orizz.inc" -->
		<!-- #include virtual="/public/layout/include/menu_vert_sx.inc" -->
		<div id="content-center">
			<!-- #include virtual="/public/layout/include/menutips.inc" -->

			<div align="left" id="contenuti">
			<form action="" method="post" name="form_field_filter" action="<%=request.ServerVariables("URL")%>">		
			<input type="hidden" value="<%=strGerarchia%>" name="gerarchia">	
			<input type="hidden" value="<%=numPage%>" name="page">
			<input type="hidden" value="<%=order_by%>" name="order_by">   
			<input type="hidden" value="0" name="reset_fields_filter" id="reset_fields_filter">  
			<input type="hidden" value="1" name="fields_filter" id="fields_filter"> 
			
			<select name="field_provincia" id="field_provincia">
			<option value=""></option>
			<%
			On Error Resume Next
			Set objListProv = objContentField.getListContentFieldValuesByDesc("provincia", langIdTarget, null)
			for each k in objListProv
				label=k
				if not(lang.getTranslated("portal.commons.content_field.label."&label)="") then label=lang.getTranslated("portal.commons.content_field.label."&label) end if%>
				<option value="<%=objListProv(k)&"#"&k%>" <%if(strComp(objListPairKeyValue(Cstr(objListProv(k))), k, 1) = 0)then response.write("selected") end if%>><%=label%></option>
			<%next
			if(Err.number<>0)then
			response.write(Err.description)
			end if%>
			</select>

			<select name="field_citta" id="field_citta">
			<option value=""></option>
			<%
			On Error Resume Next
			Set objListCity = objContentField.getListContentFieldValuesByDesc("città", langIdTarget, null)
			for each j in objListCity
				label=j
				if not(lang.getTranslated("portal.commons.content_field.label."&label)="") then label=lang.getTranslated("portal.commons.content_field.label."&label) end if%>
				<option value="<%=objListCity(j)&"#"&j%>" <%if(strComp(objListPairKeyValue(Cstr(objListCity(j))), j, 1) = 0)then response.write("selected") end if%>><%=label%></option>
			<%next
			if(Err.number<>0)then
			response.write(Err.description)
			end if%>
			</select>
			
			<input type="submit" value="<%=lang.getTranslated("frontend.template_fields.filter.label.submit")%>">&nbsp;<input type="button" value="<%=lang.getTranslated("frontend.template_fields.filter.label.reset")%>" onclick="javascript:resetFilters();">
			</form>

			<script>
			$('#field_provincia').change(function() {
				var field_provincia_val_ch = $('#field_provincia').val();
				field_provincia_val_ch = field_provincia_val_ch.substring(field_provincia_val_ch.indexOf("#")+1);
				
				$('select#field_citta').find('option').each(function() {
					$(this).show();
					var tmpVal = $(this).val();
					tmpVal = tmpVal.substring(tmpVal.indexOf("#")+1,tmpVal.indexOf("-"));
					if(tmpVal!=field_provincia_val_ch){
						$(this).hide();
					}
				});
			});			
			</script>			
			
			<%
			'************** codice per la lista news e paginazione
			if(bolHasObj) then%>
				<br/>			
				<%		
				for newsCounter = FromNews to ToNews
					Set objSelNews = objTmpNews(newsCounter)
					detailURL = "#"
					if(bolHasDetailLink) then
						detailURL = objMenuFruizione.resolveHrefUrl(base_url, (modelPageNum+1), lang, objCategoriaTmp, objTemplateSelected, objPageTempl)
					end if%>
					<div><p class="title_contenuti"><a href="javascript:openDetailContentPage('<%=detailURL%>', '<%=strGerarchia%>', <%=objSelNews.getNewsID()%>, <%=(modelPageNum+1)%>);"><%=objSelNews.getTitolo()%></a></p>
					<%if (Len(objSelNews.getAbstract1()) > 0) then response.Write(objSelNews.getAbstract1()) end if%>					
					
					
					<%if(strComp(typename(objSelNews.getListaFields()), "Dictionary", 1) = 0)then	
						
						response.write("<p>")
						Set objListContentField = objSelNews.getListaFields()		
						for each k in objListContentField
							On Error Resume next
							Set objField = objListContentField(k)
							labelForm = objField.getDescription()
							if not(lang.getTranslated("frontend.contenuto.field.label."&objField.getDescription())="") then labelForm = lang.getTranslated("frontend.contenuto.field.label."&objField.getDescription())

							'*** imposto la descrizione per il gruppo di appartenenza
							if(strComp(typename(objField.getObjGroup()), "ProductFieldGroupClass") = 0)then
								tmpDescG = objField.getObjGroup().getDescription()
								if(tmpDescG <> tmpGroupDesc)then
									tmpGroupDesc = tmpDescG
							tmpGroupDescTrans = tmpGroupDesc
									if not(lang.getTranslated("frontend.contenuto.field.label.group."&tmpGroupDesc)="") then tmpGroupDescTrans = lang.getTranslated("frontend.contenuto.field.label.group."&tmpGroupDesc)
										
									labelForm = "<div class=""contenuto_field_contenuto_group"">"& tmpGroupDescTrans & "</div>" & labelForm
								end if
							end if

							fieldCssClass=""

							select Case objField.getTypeField()								
							Case 1,2
								fieldCssClass="formFieldTXTMedium"
								if(objField.getEditable()="1")then
									response.write(labelForm & ":&nbsp;" &objContentField.renderContentFieldHTML(objField,fieldCssClass, "", objCurrentNews.getNewsID(), "",lang,1,objField.getEditable()) & "<br/>")%>
								<%
								else
									valueTmp = objField.getSelValue()
									if not(lang.getTranslated("portal.commons.content_field.label."&valueTmp)="") then valueTmp=lang.getTranslated("portal.commons.content_field.label."&valueTmp) end if
									response.write(labelForm & ":&nbsp;" & valueTmp & "<br/>")
								end if			
							Case 3,4,5,6						
								if(CInt(objField.getTypeField())=4) then
									fieldCssClass="formFieldMultiple"
								end if
								response.write(labelForm & ":&nbsp;" &objContentField.renderContentFieldHTML(objField,fieldCssClass, "", objCurrentNews.getNewsID(), "",lang,1,objField.getEditable()) & "<br/>")
							Case 7
							  fieldValueMatch = objContentField.findFieldMatchValue(k,objCurrentNews.getNewsID())
							  response.write(objContentField.renderContentFieldHTML(objField,fieldCssClass, "", objCurrentNews.getNewsID(), fieldValueMatch,lang,1,objField.getEditable()))
							Case 8
								fieldCssClass="formFieldTXTMedium"
								if(objField.getEditable()="1")then
									response.write(labelForm & ":&nbsp;" &objContentField.renderContentFieldHTML(objField,fieldCssClass, "", objCurrentNews.getNewsID(), "",lang,1,objField.getEditable()) & "<br/>")%>
								<%
								else
									valueTmp = objField.getSelValue()
									if not(lang.getTranslated("portal.commons.content_field.label."&valueTmp)="") then valueTmp=lang.getTranslated("portal.commons.content_field.label."&valueTmp) end if
									response.write(labelForm & ":&nbsp;" & valueTmp & "<br/>")
								end if	
							Case 9
								fieldCssClass="formFieldTXTMedium"
								if(objField.getEditable()="1")then%>
									<script>
									    //declare cleditor option array;
									    var cloptions<%=objContentField.getFieldPrefix()&objField.getID()%> = {
									    width:280,	// width not including margins, borders or padding
									    height:200,	// height not including margins, borders or padding
									    controls:"bold italic underline strikethrough subscript superscript | font size style | color highlight removeformat | bullets numbering | alignleft center alignright justify | rule | cut copy paste | image",	// controls to add to the toolbar
									    }
									</script>
									<%response.write(labelForm & ":&nbsp;" &objContentField.renderContentFieldHTML(objField,fieldCssClass, "", objCurrentNews.getNewsID(), "",lang,1,objField.getEditable()) & "<br/>")%>
								<%
								else
									valueTmp = Server.HTMLEncode(objField.getSelValue())
									if not(lang.getTranslated("portal.commons.content_field.label."&valueTmp)="") then valueTmp=lang.getTranslated("portal.commons.content_field.label."&valueTmp) end if
									response.write(labelForm & ":&nbsp;" & valueTmp & "<br/>")
								end if
							Case Else
							End Select

							Set objField = nothing

							if(Err.number<>0) then
							'response.write(Err.description)
							end if
						next
						
						response.write("</p>")						
						
						Set objListContentField = nothing
					end if%>					
					
					
					</div>
					<p class="line"></p>
					<%Set objSelNews = nothing
				next%>
				<div><%if(totPages > 1) then call PaginazioneFrontend(totPages, numPage, strGerarchia, request.ServerVariables("URL"), strParamPagFilter) end if%></div>
				<div id="torna"><a href="<%=Application("baseroot") & "/common/include/feedRSS.asp?gerarchia="&strGerarchia%>" target="_blank"><img src="<%=Application("baseroot")&"/common/img/rss_image.gif"%>" vspace="3" hspace="3" border="0" align="right" alt="RSS"></a></div>
			<%else%>
				<br/><br/><div align="center"><strong><%=lang.getTranslated("portal.commons.templates.label.page_in_progress")%></strong></div>
			<%end if%>
			</div>
			<form action="" method="post" name="form_detail_link_news">	
			<input type="hidden" value="" name="id_news">	
			<input type="hidden" value="" name="modelPageNum">	
			<input type="hidden" value="" name="gerarchia">	
			<input type="hidden" value="<%=numPage%>" name="page">
			<input type="hidden" value="<%=order_by%>" name="order_by">            
			</form>	
		</div>
		<!-- #include virtual="/public/layout/include/menu_vert_dx.inc" -->
	</div>
	<!-- #include virtual="/public/layout/include/bottom.inc" -->
</div>
</body>
</html>
<%
'****************************** PULIZIA DEGLI OGGETTI UTILIZZATI
Set objContentField = nothing
Set objCat = nothing
Set objPageTempl = nothing
Set objTemplate = nothing
Set objMenuFruizione = nothing
Set objListPoint = nothing
Set objListaTargetCat = nothing
Set objListaTargetLang = nothing
Set objListaNews = nothing
Set News = Nothing
%>
