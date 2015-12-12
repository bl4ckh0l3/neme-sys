<%
'***************** START: recupero utente e ruolo
Set objUserLogged = new UserClass
if not(isEmpty(Session("objUtenteLogged"))) AND not(Session("objUtenteLogged") = "") then
	Set objUserLoggedTmp = objUserLogged.findUserByID(Session("objUtenteLogged"))
	if not(isNull(objUserLoggedTmp)) AND (Instr(1, typename(objUserLoggedTmp), "UserClass", 1) > 0) then 
		strRuoloLogged = objUserLoggedTmp.getRuolo()  
		numIdUser = objUserLoggedTmp.getUserID()
	end if
	Set objUserLoggedTmp = nothing
	isLogged = true
elseif not(isEmpty(Session("objCMSUtenteLogged"))) AND not(Session("objCMSUtenteLogged") = "") then
	Set objCMSUtenteLoggedTmp = objUserLogged.findUserByID(Session("objCMSUtenteLogged"))
	if not(isNull(objCMSUtenteLoggedTmp)) AND (Instr(1, typename(objCMSUtenteLoggedTmp), "UserClass", 1) > 0) then 
		strRuoloLogged = objCMSUtenteLoggedTmp.getRuolo()
		numIdUser = objCMSUtenteLoggedTmp.getUserID() 
	end if
	Set objCMSUtenteLoggedTmp = nothing
	isLogged = true
end if
Set objUserLogged = nothing
'***************** END: recupero utente e ruolo

'***************** per il funzionamento del widget è necessario id_news (id contenuto) valorizzato  
if not(id_news="") then%>
 
<script>
  function reloadNewsCommentWidget(idNews){
	var query_string = "id_news="+idNews;

	$.ajax({
	   type: "GET",
	   cache: false,
	   url: "<%=Application("baseroot") & "/public/layout/addson/contents/news_comments_reload_widget.asp"%>",
	   data: query_string,
		success: function(html) {
		  //alert("ciao");
		  $("#ncwList").empty();
		  $("#ncwList").append(html);
		},
		error: function (xhr, ajaxOptions, thrownError){
		  //alert(xhr.status);
		  //alert(thrownError);
		}
	 });            
  }


  $(document).ready(function() {
	$(document).oneTime(30000, function() {
	  reloadNewsCommentWidget(<%=id_news%>);
	}, 1);
  });
</script>
         
  <div align="left" id="div_ncw" style="margin-top:10px;width:100%;">
        <div id="view-comments">
		<%if (isLogged) then
		  if (CInt(strRuoloLogged) = Application("admin_role")) then%>
			<a href="javascript:openWin('<%=Application("baseroot")&"/public/layout/include/popupInsertComments.asp?id_element="&id_news&"&element_type=1"%>','popupallegati',400,400,100,100);"><img alt="<%=lang.getTranslated("frontend.popup.label.insert_commento")%>" src="<%=Application("baseroot")&"/common/img/comment_add.png"%>" hspace="0" vspace="0" border="0"></a>
		  <%else%>
			<a href="javascript:prepareComment();"><img alt="<%=lang.getTranslated("frontend.popup.label.insert_commento")%>" src="<%=Application("baseroot")&"/common/img/comment_add.png"%>" hspace="0" vspace="0" border="0"></a>
		  <%end if
		else%>
				<a href="<%=Application("baseroot")&"/login.asp?from="&Application("baseroot") & "/common/include/controller.asp?"&Server.URLEncode(Request.QueryString)%>"><img alt="<%=lang.getTranslated("frontend.popup.label.insert_commento")%>" src="<%=Application("baseroot")&"/common/img/comment_add.png"%>" hspace="0" vspace="0" border="0"></a>
		<%end if%>
		<%="&nbsp;&nbsp;"&lang.getTranslated("portal.templates.commons.label.see_comments_news")%><br/>
        </div>

        <div id="comments-widget">
        <%if(request("vode_done")="1") then%>
        <span id="vote-confirmed"><%=lang.getTranslated("portal.templates.commons.label.vote_done")%></span><br/>
        <%elseif(request("vode_done")="0") then%>
        <span id="vote-confirmed"><%=lang.getTranslated("portal.templates.commons.label.vote_not_done")%></span><br/>
        <%elseif(request("add_done")="1") then%>
        <span id="vote-confirmed"><%=lang.getTranslated("portal.templates.commons.label.add_done")%></span><br/>			
        <%elseif(request("add_done")="0") then%>
        <span id="vote-confirmed"><%=lang.getTranslated("portal.templates.commons.label.add_not_done")%></span><br/>			
        <%end if%>
        <%if(request("posted")="1") then%>
          <span id="vote-confirmed">
        <%if(Application("use_comments_filter")=1) then%>
          <%=lang.getTranslated("portal.templates.commons.label.comment_posted_standby")%>
        <%else%>
          <%=lang.getTranslated("portal.templates.commons.label.comment_posted")%>
        <%end if%>
        </span><br/>
        <%elseif(request("posted")="0") then%>
          <span id="vote-confirmed"><%=lang.getTranslated("portal.templates.commons.label.comment_no_posted")%></span><br/>
        <%end if
 
        Set objCommento = New CommentsClass
        
        Dim commentsFound
        commentsFound = false
        
        on error Resume Next
        
        if not(id_news="") AND objCommento.findCommentiByIDElement(id_news,1,1).Count > 0 then
          commentsFound = true
        end if
        
        if Err.number <> 0 then
          'response.write(Err.description)
          commentsFound = false
        end if	
        
        if (commentsFound) then
          Dim  kk, objTmpCommento, objUserClass, usrHasImgComment, commentCounter		
          Set objSelectedCommento = objCommento.findCommentiByIDElement(id_news,1,1)
          Set objUserClass = new UserClass
					Set objUserPreference = new UserPreferenceClass
        
          commentCounter = 0
        
          for each kk in objSelectedCommento.Keys
            'response.write(kk&"<br>")
            objUsrCanVote = true
            Set objTmpCommento = objSelectedCommento(kk)
            Set objUserComment = objUserClass.findUserByID(objTmpCommento.getIDUtente())
            usrHasImgComment = objUserClass.hasImageUser(objTmpCommento.getIDUtente())
            likeCount = 0
            nolikeCount = 0
            on error Resume Next
            Set objLPC = objUserPreference.getListUserPreferenceByFriendAndComment(objTmpCommento.getIDUtente(), null, objTmpCommento.getIDCommento())
            for each h in objLPC
            	if(objLPC(h).getType()=1) then
                likeCount=likeCount+1
              elseif(objLPC(h).getType()=0)then
                nolikeCount=nolikeCount+1    
              end if  
            
              if (objLPC(h).getIdFriend()=numIdUser)then
                  objUsrCanVote=objUsrCanVote AND false
              end if
                
              'response.write("objLPC(h).getIdFriend(): "&objLPC(h).getIdFriend()&" - numIdUser: "&numIdUser&" - likeCount: "&likeCount&" - nolikeCount: "&nolikeCount&"<br>")
            next  
            if Err.number <> 0 then
              likeCount = 0
              nolikeCount = 0
              objUsrCanVote = true            
            end if              
            %>
            <div class="commento" id="comment_<%=commentCounter%>">
              <br/><br/>
              <div style="float:left;padding:0px 5px 5px 0px;width:50px;height:50px;overflow:hidden;text-align:center;">
              <%if (usrHasImgComment) then%>
              <img class="imgAvatarUserNCW" src="<%=Application("baseroot") & "/common/include/userImage.asp?userID="&objUserComment.getUserID()%>"/>
              <%else%>
              <img class="imgAvatarUserNCW" src="<%=Application("baseroot") & "/common/img/unkow-user.jpg"%>"/>
              <%end if%>
              </div>
              
              <div style="display:inline-block;padding:0px 5px 5px 0px;">
              <%if (isLogged) then
                if (CInt(strRuoloLogged) = Application("admin_role")) then%>
                  <a href="javascript:sendAjaxDelComment(<%=objTmpCommento.getIDCommento()%>,<%=id_news%>,'comment_<%=commentCounter%>');">x</a>&nbsp;&nbsp;
                <%end if
              end if%>
              <strong><%=objTmpCommento.getDtaInserimento()&"&nbsp;"%>
			<!--nsys-modcommunity6-->
			<%if(objUserComment.getPublic()) then%>
			      <span id="showprofilenc<%=commentCounter%>_<%=objTmpCommento.getIDUtente()%>"><a title="<%=lang.getTranslated("portal.templates.commons.label.view_pub_profile")%>" href="<%=Application("baseroot") & "/area_user/publicprofile.asp?id_utente="&objTmpCommento.getIDUtente()%>"><%=objUserComment.getUsername()%></a></span>
			      <span id="shownamenc<%=commentCounter%>_<%=objTmpCommento.getIDUtente()%>"></span><br/> 
				<%
			      objUsrCanVote=objUsrCanVote AND (numIdUser<>"" AND numIdUser <>"-1" AND objTmpCommento.getIDUtente()<>numIdUser)%>
			      <span id="showlikenc<%=commentCounter%>_<%=objTmpCommento.getIDUtente()%>"><a class="addcommentvote<%=commentCounter%>" href="javascript:prepareVote(<%=objUserComment.getUserID()%>, <%=objTmpCommento.getIDCommento()%>, 1, 1);"><img id="ok<%=commentCounter%>" title="<%=lang.getTranslated("portal.templates.commons.label.vote_up")%>" src="<%=Application("baseroot") & "/common/img/vote_ok.png"%>"/></a><%if(likeCount>0)then response.write("<span class=likecountpref>"&likeCount&"</span>") end if%>
			      &nbsp;&nbsp;<a class="addcommentvote<%=commentCounter%>" href="javascript:prepareVote(<%=objUserComment.getUserID()%>, <%=objTmpCommento.getIDCommento()%>, 0, 1);"><img id="ko<%=commentCounter%>" title="<%=lang.getTranslated("portal.templates.commons.label.vote_down")%>" src="<%=Application("baseroot") & "/common/img/vote_ko.png"%>"/></a><%if(nolikeCount>0)then response.write("<span class=likecountpref>"&nolikeCount&"</span>") end if%></span>

				<span id="showaddfnc<%=commentCounter%>_<%=objTmpCommento.getIDUtente()%>">&nbsp;&nbsp;<a href="javascript:addFriend(<%=objUserComment.getUserID()%>);"><img id="add<%=commentCounter%>" title="<%=lang.getTranslated("portal.templates.commons.label.add_friend")%>" src="<%=Application("baseroot") & "/common/img/group_link.png"%>"/></a></span>
			      <script>
				$("#showaddfnc<%=commentCounter%>_<%=objTmpCommento.getIDUtente()%>").hide();                  
			      checkAjaxHasFriendNC('showaddfnc<%=commentCounter%>_',<%=objTmpCommento.getIDUtente()%>, '<%=objUserComment.getUsername()%>');
	
				<%if not(objUsrCanVote)then%> 
				$("a.addcommentvote<%=commentCounter%>").attr('href', 'javascript:alert("<%=lang.getTranslated("portal.templates.commons.label.vote_cannot_done")%>")');   
				$("a.addcommentvote<%=commentCounter%>").attr('style', 'cursor: default');               
				<%end if%>		
						$("#showprofilenc<%=commentCounter%>_<%=objTmpCommento.getIDUtente()%>").hide();                  
			      checkAjaxHasFriendActiveNC('showprofilenc<%=commentCounter%>_','shownamenc<%=commentCounter%>_',<%=objTmpCommento.getIDUtente()%>, '<%=objUserComment.getUsername()%>');
			      </script>
			    <%else%>
				<%=objUserComment.getUsername()%>
			    <%end if%> 
			<!---nsys-modcommunity6-->
              </strong>             
              <br><%if(objTmpCommento.getVoteType()=1)then%><img id="nolike<%=commentCounter%>" src="<%=Application("baseroot") & "/common/img/like.png"%>" align="absbottom"/><%else%><img id="nolike<%=commentCounter%>" src="<%=Application("baseroot") & "/common/img/nolike.png"%>" align="absbottom"/><%end if%>&nbsp;<%=objTmpCommento.getMessage()%><br>
              </div>
            </div>
            
            <%Set objUserComment = nothing
            	Set objLPC = nothing
              commentCounter = commentCounter+1
          next%>
          <script>
           $(function() {
            $('.imgAvatarUserNCW').aeImageResize({height: 50, width: 50});
          });   
          </script>              
          <%Set objUserClass = nothing
          Set objUserPreference = nothing
        else
          response.Write("<br/><span align='center'>"&lang.getTranslated("frontend.popup.label.no_comments_news")&"</span><br>")
        end if				
        Set objCommento = Nothing
        %>
        </div>
</div>        	
<%end if%>