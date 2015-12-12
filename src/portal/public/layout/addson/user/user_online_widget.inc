<div id="onlineUsersList">
	<script>
			function reloadUsrOnline(){
		  $.ajax({
			 type: "GET",
			 cache: false,
			 url: "<%=Application("baseroot") & "/public/layout/addson/user/user_online_widget.asp"%>",
			  success: function(html) {
					//alert("ciao");
					$("#onlineUsersList").empty();
					$("#onlineUsersList").append(html);
			  },
			  error: function (xhr, ajaxOptions, thrownError){
										//alert(xhr.status);
										//alert(thrownError);
									}
		   });            
			}
	
	
						$(document).ready(function() {
							$(document).oneTime(30000, function() {
								reloadUsrOnline();
							}, 1);
						});
		  
		function addAjaxFriend(id_friend, active){
		  var query_string = "id_utente="+id_friend+"&active="+active+"&action=1";
		
		  $.ajax({
			 type: "POST",
			 cache: false,
			 url: "<%=Application("baseroot") & "/area_user/checkfriend.asp"%>",
			 data: query_string,
			  success: function(response) {
				// show friend request icon
				//alert("response: "+response);
				if(response==1){
				$("#addfriend_"+id_friend).hide();
				}
			  },
			  error: function() {
				//$("#addfriend_"+id_friend).show();
			  }
		   });
		}

		function checkAjaxHasFriend(id_friend, active){
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
				$("#addfriend_"+id_friend).show();
				}
			  },
			  error: function() {
				$("#addfriend_"+id_friend).hide();
			  }
		   });
		}

		function checkAjaxHasFriendActive(id_friend, usrnameCurrUser){
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
										$("#showprofilew_"+id_friend).hide();
										$("#shownamew_"+id_friend).empty().append(usrnameCurrUser);
										}else{
										$("#shownamew_"+id_friend).empty();
										$("#showprofilew_"+id_friend).show();					
										}
			  },
			  error: function() {
										$("#showprofilew_"+id_friend).hide();
										$("#shownamew_"+id_friend).empty().append(usrnameCurrUser);
			  }
		   });
		}
	</script>
      <%if(onlineUsersList.Count>0) then%>
        <div id="online-users" style="height:200px;overflow:auto;">
      
      
        <h2><%=lang.getTranslated("frontend.menu.label.online_users_list")%></h2>
          <%for each x in onlineUsersList
              arrUsrInfo = Split(onlineUsersList(x),"|")
              idCurrUser = arrUsrInfo(0)
              hasUAv = arrUsrInfo(2)
              usrnameCurrUser = arrUsrInfo(3)%>
              <div style="float:left;width:50px;height:50px;overflow:hidden;text-align:center;padding-top:1px;padding-right:2px;">
              <%if(hasUAv)then%>
                <img class="imgAvatarUserOn" src="<%=Application("baseroot") & "/common/include/userImage.asp?userID="&idCurrUser%>" <%If (InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE",1) > 0) then response.write(" width=""50"" height=""50""") end if%> />
              <%else%>
                <img class="imgAvatarUserOn" src="<%=Application("baseroot") & "/common/img/unkow-user.jpg"%>" <%If (InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE",1) > 0) then response.write(" width=""50"" height=""50""") end if%> />
              <%end if%>
              </div> 
              <div style="margin-bottom:40px;border-top:1px solid;height:12px;text-align:left;font-size:12px;">
              <!--nsys-modcommunity5-->
              <%if(arrUsrInfo(1)=1) then%>
                <span id="showprofilew_<%=idCurrUser%>"><a title="<%=lang.getTranslated("portal.templates.commons.label.view_pub_profile")%>" href="<%=Application("baseroot") & "/area_user/publicprofile.asp?id_utente="&idCurrUser%>"><%=usrnameCurrUser%></a></span>
                <span id="shownamew_<%=idCurrUser%>"></span>
                <script>
								$("#showprofilew_<%=idCurrUser%>").hide();                  
                checkAjaxHasFriendActive(<%=idCurrUser%>, '<%=usrnameCurrUser%>');
                </script>
              <%else%>
                <%=usrnameCurrUser%>
              <%end if%>
              <%
              percentualO = 0
              Set objUserPreferenceO = new UserPreferenceClass
              percentualO = objUserPreferenceO.findUserPreferencePositivePercent(idCurrUser)
              percentualO=FormatNumber(percentualO, 0,-1)              
              endcounter=0      
              if(percentualO>0 AND percentualO<=20)then
                endcounter=1
              elseif(percentualO>20 AND percentualO<=40)then 
                endcounter=2
              elseif(percentualO>40 AND percentualO<=60)then
                endcounter=3
              elseif(percentualO>60 AND percentualO<=80)then 
                endcounter=4
              elseif(percentualO>80 AND percentualO<=100)then 
                endcounter=5
              end if
              if(endcounter>0)then
              %>
              <br/>
              <%for starcount = 1 to endcounter%>
              <img width="14" height="15" src="<%=Application("baseroot")&"/common/img/ico_stella.png"%>" align="absmiddle" style="padding:0px;border:0px;">
              <%next%>
              <%end if
              Set objUserPreferenceO = nothing
              %>              
              <!---nsys-modcommunity5-->
                <!--nsys-modcommunity4-->
                <%if(arrUsrInfo(1)=1) then%>  
                <br/><a href="javascript:addAjaxFriend(<%=idCurrUser%>,0);" title="<%=lang.getTranslated("portal.templates.commons.label.add_friend")%>"><img id="addfriend_<%=idCurrUser%>" alt="<%=lang.getTranslated("portal.templates.commons.label.add_friend")%>" src="<%=Application("baseroot") & "/common/img/group_link.png"%>"/></a>
                <script>
                $('#addfriend_<%=idCurrUser%>').hide();                
                checkAjaxHasFriend(<%=idCurrUser%>, 0);
                </script>
                <%end if%>
                <!---nsys-modcommunity4-->
              </div>           
          <%next%>
        <script>
          $(function() {
            $('.imgAvatarUserOn').aeImageResize({height: 50, width: 50});
          });
        </script>
        </div>
      <%end if%>
</div>