<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<!-- #include virtual="/common/include/Objects/ContentFieldGroupClass.asp" -->
<!-- #include virtual="/common/include/Objects/ContentFieldClass.asp" -->
<!-- #include virtual="/common/include/Paginazione.inc" -->
<!-- #include file="include/init2.inc" -->
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
</head>
<body>
<div id="warp">
	<!-- #include virtual="/public/layout/include/header.inc" -->	
	<div id="container">	
		<!-- include virtual="/public/layout/include/menu_orizz.inc" -->
		<!-- #include virtual="/public/layout/include/menu_vert_sx.inc" -->
		<div id="content-center">
			<!-- #include virtual="/public/layout/include/menutips.inc" -->

			<div align="left">
			<%if (bolHasObj) then%>
				<div>
				<p><strong><%=objCurrentNews.getTitolo()%></strong></p>
				<%
				if (Len(objCurrentNews.getAbstract1()) > 0) then response.Write(objCurrentNews.getAbstract1()) end if
				if (Len(objCurrentNews.getAbstract2()) > 0) then response.Write(objCurrentNews.getAbstract2()) end if
				if (Len(objCurrentNews.getAbstract3()) > 0) then response.Write(objCurrentNews.getAbstract3()) end if
				response.Write(objCurrentNews.getTesto())


				On Error Resume Next
				hasContentFields=false
				
				Set objListContentField = objContentField.getListContentField4ContentActive(objCurrentNews.getNewsID())
				
				if(objListContentField.Count > 0)then
					hasContentFields = true
				end if
				
				if(Err.number <> 0) then
					hasContentFields = false
				end if	
				
				if(hasContentFields)then	
					
					response.write("<p>")
								
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
				end if



				if(bolHasAttach) then 
					for each key in attachMap
						if(attachMap(key).count > 0)then%>
							<br/><br/><strong><%=lang.getTranslated(attachMultiLangKey(key))%></strong><br/>
							<%for each item in attachMap(key)%>
								<a href="javascript:openWin('<%=Application("baseroot")&"/public/layout/include/popup.asp?id_allegato="&item.getFileID()&"&parent_type=1"%>','popupallegati',400,400,100,100)"><%=item.getFileName()%></a><br>
							<%next
						end if
					next
				end if
				Set objCurrentNews = nothing				
				%>
				</div>
				<div id="torna"><a href="<%=Application("baseroot") & "/common/include/feedRSS.asp?gerarchia="&strGerarchia&"&id_news="&id_news&"&page="&numPage&"&modelPageNum="&modelPageNum%>" target="_blank"><img src="<%=Application("baseroot")&"/common/img/rss_image.gif"%>" vspace="3" hspace="3" border="0" align="right" alt="RSS"></a></div>
			<%else%>
				<br/><br/><div align="center"><strong><%=lang.getTranslated("portal.commons.templates.label.page_in_progress")%></strong></div>
			<%end if%>
			</div>
		</div>
		<!-- #include virtual="/public/layout/include/menu_vert_dx.inc" -->
		<!-- #include virtual="/public/layout/addson/contents/news_comments_widget.inc" -->
	</div>
	<!-- #include virtual="/public/layout/include/bottom.inc" -->
</div>
</body>
</html>
<%
'****************************** PULIZIA DEGLI OGGETTI UTILIZZATI
Set objContentField = nothing
Set objContentFieldGroup = nothing
Set objListaTargetCat = nothing
Set objListaTargetLang = nothing
Set objListaNews = nothing
Set News = Nothing
%>