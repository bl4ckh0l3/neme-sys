

 $(document).ready(function(){

	$(".fix_notes_vedi").click(function(event) {
		$(this).next().slideToggle('hide');
		 event.preventDefault();

		return false;
	
	});
	
	
	
	
	$(".quantita_meno").click(function(event) {
		var n = Number($(this).next('input').val());
		var str = (n - 1);
	
		if(str>-1){	
			$(this).next('input').val(str);
		}
			
		event.preventDefault();
		return false;	
	});
	
	
	$(".quantita_plus").click(function(event) {
		var n = Number($(this).prev('input').val());
		var str = ( n + 1);

		$(this).prev('input').val(str);
			

		event.preventDefault();
		return false;	
	});
});




function doSearch(){
	if(document.search_news.search_full_txt.value == ""){
		alert("Inserire un valore per la ricerca!");
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




function sendLoginForm(){
					if(document.login.j_username.value == "USERNAME"){
						document.login.j_username.value = "";
					}
					if(document.login.j_username.value == ""){
						alert("Inserire nome utente!");
						document.login.j_username.focus();
						return false;						
					}
					
					if(document.login.j_password.value == ""){
						alert("Inserire password");
						if(document.getElementById('divpwd2').style.display=="visible"){
							document.login.j_password.focus();
						}
						return false;
					}					
					
					document.login.submit();
				}
				
function cleanLoginField(formfieldId){
	var elem = document.getElementById(formfieldId);
	elem.value="";
}
				
				function restoreLoginField(formfieldId, valueField){
				  var elem = document.getElementById(formfieldId);
				  if(elem.value==''){
					elem.value=valueField;
				  }
				}
				