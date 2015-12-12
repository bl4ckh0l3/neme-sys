<link rel="stylesheet" href="<%=Application("baseroot") & "/editor/css/jquery-ui-latest.custom.css"%>" type="text/css">
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery-latest.min.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery-ui-latest.custom.min.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery.google-analytics.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery-ui-timepicker-addon.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery.form.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/javascript_global.js"%>"></script>

<!-- carico l'editor html semplificato CLEditor -->
<link rel="stylesheet" type="text/css" href="<%=Application("baseroot")&"/cleditor/jquery.cleditor.css"%>" />      
<script type="text/javascript" src="<%=Application("baseroot")&"/cleditor/jquery.cleditor.js"%>"></script>

<!-- gestione degli update fields via ajax -->
<script language="JavaScript">
function sendAjaxCommand(field_name, field_val, objtype, id_objref, listCounter, field){
	//var query_string = "field_name="+field_name+"&field_val="+escape(field_val)+"&objtype="+objtype+"&id_objref="+id_objref;
	var query_string = "field_name="+field_name+"&field_val="+encodeURIComponent(field_val)+"&objtype="+objtype+"&id_objref="+id_objref;
	//alert("query_string: "+query_string);
	var resp = false;

	$.ajax({
		async: false,
		type: "GET",
		cache: false,
		url: "<%=Application("baseroot") & "/editor/include/ajaxupdate.asp"%>",
		data: query_string,
		success: function(response) {
			//alert("response: "+response);
			/*$("#ajaxresp").empty();
			$("#ajaxresp").append("<%=langEditor.getTranslated("backend.commons.ok_updated_field")%>");
			$("#ajaxresp").fadeIn(1500,"linear");
			$("#ajaxresp").fadeOut(600,"linear");*/
			resp = true;

			// il codice seguente server per inviare il contatore dell'oggetto modificato nella lista
			// per chiamare la funzione specifica di ogni pagina, per modificare elementi della pagina accessori
			if(typeof changeRowListData == 'function'){				
				changeRowListData(listCounter, objtype, field);
			}
		},
		error: function() {
			$("#ajaxresp").empty();
			$("#ajaxresp").append("<%=langEditor.getTranslated("backend.commons.fail_updated_field")%>");
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
		url: "<%=Application("baseroot") & "/editor/include/ajaxdelete.asp"%>",
		data: query_string,
		success: function(response) {
			if(response.indexOf("err:")<0){
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
        response = response.replace("err:","");
				location.href='<%=Application("baseroot")&Application("error_page")&"?error="%>'+response;				
			}
		},
		error: function() {
			$("#ajaxresp").empty();
			$("#ajaxresp").append("<%=langEditor.getTranslated("backend.commons.fail_delete_item")%>");
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
      
<%if(Trim(Application("googlemaps_key"))<>"")then%>
<!--  ****************************************** INTEGRAZIONE GOOGLEMAP API ****************************************** -->
<script src="https://maps.googleapis.com/maps/api/js?key=<%=Trim(Application("googlemaps_key"))%>&amp;sensor=false" type="text/javascript"></script>
 <%end if%>