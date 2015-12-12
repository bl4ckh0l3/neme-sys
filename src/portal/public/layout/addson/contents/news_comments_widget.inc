 <form action="<%=Application("baseroot") & "/common/include/Controller.asp"%>" method="get" name="form_reload_page">
  <input type="hidden" name="gerarchia" value="<%=request("gerarchia")%>">
  <input type="hidden" name="id_news" value="<%=id_news%>">
  <input type="hidden" name="page" value="<%=request("page")%>">
  <input type="hidden" name="modelPageNum" value="<%=request("modelPageNum")%>">
  </form>

<script language="JavaScript">

  var commentWidgetX = 0;
  var commentWidgetY = 0;

  jQuery(document).ready(function(){
	 $(document).mousemove(function(e){
	  commentWidgetX = e.pageX;
	  commentWidgetY = e.pageY;
	 }); 
  })

  function prepareComment(){		
	var divcomment = document.getElementById("send-comment");
	var offsetx   = 400;
	var offsety   = 50;	
	  
	if(ie||mac_ie){
	divcomment.style.left=commentWidgetX-offsetx;
	divcomment.style.top=commentWidgetY-offsety;
	}else{
	divcomment.style.left=commentWidgetX-offsetx+"px";
	divcomment.style.top=commentWidgetY-offsety+"px";
	}

	$("#send-comment").show(1000);
	divcomment.style.visibility = "visible";
	divcomment.style.display = "block";
  }

  function sendForm(){    
    if(document.form_comment.comment_message.value == ""){
      alert("<%=lang.getTranslated("frontend.popup.js.alert.insert_commento")%>");
      return;
    }else{
      document.form_comment.submit();	
    }
  }

  function prepareVote(id_utente, id_usr_comment, vote, comment_type){		
	var divvote = document.getElementById("send-vote");
	var offsetx   = 150;
	var offsety   = 150;	
	  
	if(ie||mac_ie){
	divvote.style.left=commentWidgetX-offsetx;
	divvote.style.top=commentWidgetY-offsety;
	}else{
	divvote.style.left=commentWidgetX-offsetx+"px";
	divvote.style.top=commentWidgetY-offsety+"px";
	}

	$("#send-vote").show(1000);
	divvote.style.visibility = "visible";
	divvote.style.display = "block";
	
	document.form_vote.id_utente.value=id_utente;
	document.form_vote.id_usr_comment.value=id_usr_comment;
	document.form_vote.comment_type.value=comment_type;
	document.form_vote.vote.value=vote;
	/** con le sei righe seguenti disabilito la comparsa del popup per il voto
	 immetto un messaggio standard = "a xx piace questo elemento" oppure "a xx non piace questo elemento"
	 per riabilitare l'inserimento di un messaggio da parte dell'utente oltre al voto commentare le sei righe seguenti
	 e decommentare il blocco precedente dello script **/
	/*if(vote==1){
		document.form_vote.vote_message.value = "<%=lang.getTranslated("frontend.area_user.manage.label.likemsg")%>";
	}else{
		document.form_vote.vote_message.value = "<%=lang.getTranslated("frontend.area_user.manage.label.nolikemsg")%>";
	}
	insertVote();*/
  }
  
  function insertVote(){
	$("#send-vote").hide();
	
	if(document.form_vote.vote_message.value==""){
	  if(document.form_vote.vote.value==1){
		document.form_vote.vote_message.value = "<%=lang.getTranslated("frontend.area_user.manage.label.likemsg")%>";
	  }else{
		document.form_vote.vote_message.value = "<%=lang.getTranslated("frontend.area_user.manage.label.nolikemsg")%>";
	  }
	}
	
	document.form_vote.submit();      
  }

  function addFriend(idfriend){
	if(confirm("<%=lang.getTranslated("frontend.area_user.manage.label.addthis")%>")){
	  location.href='<%=Application("baseroot") & "/area_user/addfriend.asp?id_utente="%>'+idfriend+'<%="&gerarchia="&request("gerarchia")&"&id_news="&id_news&"&page="&request("page")&"&modelPageNum="&request("modelPageNum")%>';
	}
  }
  
  function hideVoteform(){
	var divvote = document.getElementById("send-vote");
	divvote.style.visibility = "hidden";
	divvote.style.display = "none";
  }
  
  function hideCommentform(){
	var divcomment = document.getElementById("send-comment");
	divcomment.style.visibility = "hidden";
	divcomment.style.display = "none";
  }          
  
  function sendAjaxDelComment(id_comment, id_element,id_div_comment){
	var query_string = "element_type=1&del_commento=1&id_commento="+id_comment+"&id_element="+id_element;
  
	$.ajax({
	   type: "POST",
	   url: "<%=Application("baseroot") & "/public/layout/include/popupInsertComments.asp"%>",
	   data: query_string,
		success: function() {
		  // update comments-widget element
		  $("#"+id_div_comment).remove();
		  //doRefreshOfCommentsWidgetDiv();
		}
	 });
   }
  

  function checkAjaxHasFriendNC(divprofile, id_friend, active){
	var query_string = "id_utente="+id_friend+"&active="+active+"&action=0";
  
	$.ajax({
	   type: "POST",
	   cache: false,
	   url: "<%=Application("baseroot") & "/area_user/checkfriend.asp"%>",
	   data: query_string,
		success: function(response) {
		  // show friend request icon
		  //alert("response: "+response);
		  if(response==0){
		  $("#"+divprofile+id_friend).show();
		  }
		},
		error: function() {
		  $("#"+divprofile+id_friend).hide();
		}
	 });
  }

function checkAjaxHasFriendActiveNC(divprofile, divname, id_friend, usrnameCurrUser){
  var query_string = "id_utente="+id_friend+"&action=2";

  $.ajax({
	 type: "POST",
	 cache: false,
	 url: "<%=Application("baseroot") & "/area_user/checkfriend.asp"%>",
	 data: query_string,
	  success: function(response) {
		// show friend request icon
		//alert("response: "+response);
		if(response!=1){
				$("#"+divprofile+id_friend).hide();
				$("#"+divname+id_friend).empty().append(usrnameCurrUser);
				}else{
				$("#"+divname+id_friend).empty();
				$("#"+divprofile+id_friend).show();					
				}
	  },
	  error: function() {
				$(""+divprofile+id_friend).hide();
				$("#"+divname+id_friend).empty().append(usrnameCurrUser);
	  }
   });
}
        

$(function() {
	$("#send-comment").draggable();
});

$(function() {
	$("#send-vote").draggable();
});
  </script>
<div id="send-vote" style="position:absolute;left:0px;top:0px;margin-bottom:3px;vertical-align:middle;text-align:center;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:15px;background:#FFFFFF;width:320px;">
		<form action="<%=Application("baseroot") & "/area_user/processuservote.asp?gerarchia="&request("gerarchia")&"&id_news="&id_news&"&page="&request("page")&"&modelPageNum="&request("modelPageNum")%>" method="post" name="form_vote" accept-charset="UTF-8">		  
		<input type="hidden" value="" name="id_utente">
		<input type="hidden" name="vote">
		<input type="hidden" name="id_usr_comment">
		<input type="hidden" name="comment_type">
		<p align="right"><a href="javascript:hideVoteform();">x</a></p>
		<strong><%=lang.getTranslated("portal.templates.commons.label.insert_vote")%></strong><br/>
    	<textarea class="formFieldTXTTextareaComment" name="vote_message" id="vote-message" onclick="$('#vote-message').focus();"></textarea>
    &nbsp;<input name="send" align="middle" value="<%=lang.getTranslated("frontend.area_user.manage.label.do_vote")%>" type="button" onclick="javascript:insertVote();">
    </form>
</div>
<div id="send-comment" style="position:absolute;left:-0px;top:0px;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;">
	  <form action="<%=Application("baseroot") & "/area_user/processusercomment.asp?gerarchia="&request("gerarchia")&"&id_news="&id_news&"&page="&request("page")&"&modelPageNum="&request("modelPageNum")%>" method="post" name="form_comment" accept-charset="UTF-8">		  
	  <input type="hidden" name="id_element" value="<%=id_news%>">
	  <input type="hidden" name="element_type" value="1">
	  <input type="hidden" name="active" value="<%if(Application("use_comments_filter")=1) then response.write("0") else response.write("1") end if%>">
	
	  <p align="right"><a href="javascript:hideCommentform();">x</a></p>
	  
	  <div style="float:top;"><span class="labelForm"><%=lang.getTranslated("frontend.popup.label.insert_commento")%></span><br>
	  <textarea class="formFieldTXTTextareaComment" name="comment_message" id="comment_message" onclick="$('#comment_message').focus();"></textarea>
    </div> 
	  <div><span><%=lang.getTranslated("frontend.area_user.manage.label.like")%></span><br>
	  <select name="comment_type" id="comment_type">
		<OPTION VALUE="1" <%if (strComp("1", bolPublic, 1) = 0) then response.Write("selected")%>><%=lang.getTranslated("portal.commons.yes")%></OPTION>
		<OPTION VALUE="0" <%if (strComp("0", bolPublic, 1) = 0) then response.Write("selected")%>><%=lang.getTranslated("portal.commons.no")%></OPTION>
	  </select>&nbsp;&nbsp;	
	  <input type="button" name="send" style="margin-left:70px;" value="<%=lang.getTranslated("frontend.popup.label.insert_commento")%>" onclick="javascript:sendForm();">		
	  </div>
	  </form>
</div>

<div id="ncwList">
<!-- #include file="news_comments_widget.asp" -->
</div>