
			<script language="JavaScript">
				function doSearch(){
					if(document.search_news.search_full_txt.value == "" || document.search_news.search_full_txt.value == "<%=lang.getTranslated("frontend.header.label.search")%>"){
						alert("<%=lang.getTranslated("frontend.menu.js.alert.insert_search_key")%>");
						return false;
					}
					document.search_news.submit();
				}
				function cleanSearchField(formfieldId){
				  var elem = document.getElementById(formfieldId);
				  elem.value="";
				}
				
				function restoreSearchField(formfieldId, valueField){
				  var elem = document.getElementById(formfieldId);
				  if(elem.value==''){
					elem.value=valueField;
				  }
				}
			</script>
			<form method="post" name="search_news" action="<%=base_url&Application("srt_default_server_name")&Application("baseroot")&Application("dir_upload_templ")&"search/"&lang.getLangcode()&"/result_search_news.asp"%>" onSubmit="return doSearch();">
			<input name="send" align="absmiddle" class="buttonSearchHead" type="image" hspace="0" vspace="0" src="<%=Application("baseroot") & "/common/img/zoom.png"%>"><input name="search_full_txt" id="search_full_txt" type="text" value="<%=lang.getTranslated("frontend.header.label.search")%>" onfocus="cleanSearchField('search_full_txt');" onBlur="restoreSearchField('search_full_txt','<%=lang.getTranslated("frontend.header.label.search")%>');"></form>