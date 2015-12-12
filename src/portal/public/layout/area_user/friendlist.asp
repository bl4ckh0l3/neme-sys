<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/area_user.css"%>" type="text/css">
<script language="JavaScript">

function delFriend(idfriend){
	if(confirm("<%=lang.getTranslated("frontend.area_user.manage.label.delthis")%>")){
		location.href='<%=Application("baseroot") & "/area_user/delfriend.asp?id_utente="%>'+idfriend;
	}
}

function confirmFriend(idfriend, active){
	if(active==1){
		if(confirm("<%=lang.getTranslated("frontend.area_user.manage.label.confthis")%>")){
			location.href='<%=Application("baseroot") & "/area_user/confriend.asp?active="%>'+active+'&id_utente='+idfriend;
		}	
	}else{
		if(confirm("<%=lang.getTranslated("frontend.area_user.manage.label.deconfthis")%>")){
			location.href='<%=Application("baseroot") & "/area_user/confriend.asp?active="%>'+active+'&id_utente='+idfriend;
		}
	}
}
function changeTab(number){
	if(number==1)
		location.href='<%=Application("baseroot") & "/area_user/userprofile.asp"%>';
	else if(number==2)
		location.href='<%=Application("baseroot") & "/area_user/manageuser.asp"%>';
	else if(number==3)
		location.href='<%=Application("baseroot") & "/area_user/friendlist.asp"%>';
	else if(number==4)
		location.href='<%=Application("baseroot") & "/area_user/userphotos.asp"%>';

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
				$("#showprofile_"+id_friend).hide();
				$("#showname_"+id_friend).append(usrnameCurrUser);
				}else{
				$("#showname_"+id_friend).empty();
				$("#showprofile_"+id_friend).show();					
				}
      },
      error: function() {
				$("#showprofile_"+id_friend).hide();
				$("#showname_"+id_friend).append(usrnameCurrUser);
      }
   });
}
</script>
</head>
<body>
<!-- #include file="grid_top.asp" -->

			<h1><%=lang.getTranslated("frontend.header.label.utente_friends")%>&nbsp;<em><%=strUserName%></em></h1>
			<p>
			<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.profile")%>" type="button" onclick="javascript:changeTab(1);">
			<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.modify")%>" type="button" onclick="javascript:changeTab(2);">
			<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.friends")%>" type="button" onclick="javascript:changeTab(3);">
			<input name="profile" align="left" value="<%=lang.getTranslated("frontend.area_user.manage.label.photos")%>" type="button" onclick="javascript:changeTab(4);">
			</p>
			<div id="profilo-utente">        
			<br/>
			<table border="0" cellpadding="0" cellspacing="0" class="principal">
				<tr> 
				<th>&nbsp;</th>
				<th align="center" width="25">&nbsp;</th>
				<th><%=lang.getTranslated("frontend.area_user.manage.label.username")%></th>
				<th><%=lang.getTranslated("frontend.area_user.manage.label.public_profile")%></th>
				<th><%=lang.getTranslated("frontend.area_user.manage.label.dta_inser")%></th>
				<th><%=lang.getTranslated("frontend.area_user.manage.label.status")%></th>
				</tr>
				<%
				Dim hasFriend
				hasFriend = false
				on error Resume Next
				Set objUser = new UserClass
				Set objListaFriend = objUser.getListaFriends(id_utente)	
	
				if(objListaFriend.Count > 0) then
					hasFriend = true
				end if
					
				if Err.number <> 0 then
				end if	
					
				if(hasFriend) then
					intCount = 0
					
					iIndex = objListaFriend.Count
					
					FromFriend = ((numPage * friendXpage) - friendXpage)
					Diff = (iIndex - ((numPage * friendXpage)-1))
					if(Diff < 1) then
						Diff = 1
					end if
					
					ToFriend = iIndex - Diff
					
					totPages = iIndex\friendXpage
					if(totPages < 1) then
						totPages = 1
					elseif((iIndex MOD friendXpage <> 0) AND not ((totPages * friendXpage) >= iIndex)) then
						totPages = totPages +1	
					end if		
							
					objTmpFriend = objListaFriend.Items				
	
					styleRow2 = "table-list-on"					
			
					for friendCounter = FromFriend to ToFriend
						styleRow = "table-list-off"
						if(friendCounter MOD 2 = 0) then styleRow = styleRow2 end if
						Set objFilteredFriend = objTmpFriend(friendCounter)
			  			usrHasImg = objUser.hasImageUser(objFilteredFriend.getUserID())
						%>
						<tr class="<%=styleRow%>">
						<td><a title="<%=lang.getTranslated("portal.templates.commons.label.del_friend")%>" href="javascript:delFriend(<%=objFilteredFriend.getUserID()%>);"><img id="add" src="<%=Application("baseroot") & "/common/img/cancel.png"%>"/></a>
						</td>
						<td align="center">
					  <script>
						$(function() {
						  $(".imgAvatarUser").aeImageResize({height: 50, width: 50});
						});
					  </script>          
					  <%if (usrHasImg) then%>
						<img class="imgAvatarUser" src="<%=Application("baseroot") & "/common/include/userImage.asp?userID="&objFilteredFriend.getUserID()%>"/>
						<!--<script>resizeimagesByID('imgUser', 50);</script>-->
						<%else%>
						<img class="imgAvatarUser" src="<%=Application("baseroot") & "/common/img/unkow-user.jpg"%>"/>
						<%end if%></td>
						<td>
						<%if(objFilteredFriend.getPublic()) then%>
						<span id="showprofile_<%=objFilteredFriend.getUserID()%>"><a title="<%=lang.getTranslated("portal.templates.commons.label.view_pub_profile")%>" href="<%=Application("baseroot") & "/area_user/publicprofile.asp?id_utente="&objFilteredFriend.getUserID()&"&gerarchia="&request("gerarchia")%>"><%=objFilteredFriend.getUsername()%></a></span>
						<span id="showname_<%=objFilteredFriend.getUserID()%>"></span>
						<script>
						$("#showprofile_<%=objFilteredFriend.getUserID()%>").hide();           
						checkAjaxHasFriendActive(<%=objFilteredFriend.getUserID()%>, '<%=objFilteredFriend.getUsername()%>');
						</script>
						<%else%>
						<%=objFilteredFriend.getUsername()%>
						<%end if%></td>
						<td><%if(objFilteredFriend.getPublic()) then response.write(lang.getTranslated("portal.commons.yes")) else response.write(lang.getTranslated("portal.commons.no")) end if%></td>
						<td><%=FormatDateTime(objFilteredFriend.getInsertDate(),2)%></td>
						<td>
						<%if(objFilteredFriend.getFriendActive()=1) then						
							if(objUser.bolHasFriendActive(id_utente, objFilteredFriend.getUserID())) then%>
							<a title="<%=lang.getTranslated("portal.templates.commons.label.deconf_friend")%>" href="javascript:confirmFriend(<%=objFilteredFriend.getUserID()%>,0);"><img id="deconf" src="<%=Application("baseroot") & "/common/img/link.png"%>"/></a>
							<%else%>
							<img id="waitfriend" src="<%=Application("baseroot") & "/common/img/clock.png"%>" title="<%=lang.getTranslated("portal.templates.commons.label.wait_friend")%>" alt="<%=lang.getTranslated("portal.templates.commons.label.wait_friend")%>"/>
							<%end if%>
						<%else						
							if(objUser.bolHasFriendActive(id_utente, objFilteredFriend.getUserID())) then%>
								<a title="<%=lang.getTranslated("portal.templates.commons.label.confirm_friend")%>" href="javascript:confirmFriend(<%=objFilteredFriend.getUserID()%>,1);"><img id="waitconf" src="<%=Application("baseroot") & "/common/img/link_error.png"%>"/></a>
							<%else%>
								<a title="<%=lang.getTranslated("portal.templates.commons.label.conf_friend")%>" href="javascript:confirmFriend(<%=objFilteredFriend.getUserID()%>,1);"><img id="conf" src="<%=Application("baseroot") & "/common/img/link_break.png"%>"/></a>							
							<%end if%>
						<%end if%>&nbsp;</td>
						</tr>				
						<%intCount = intCount +1
						Set objFilteredFriend = nothing
					next
					Set objTmpFriend = nothing
					Set objListaFriend = nothing
					Set objUser = nothing
					%>
				  
					<tr> 
					<th colspan="6" align="center">
					<%		
					'**************** richiamo paginazione
					call PaginazioneFrontend(totPages, numPage, strGerarchia, "/area_user/friendlist.asp", "order_by="&order_friend_by)%>
					</th>			
					</tr>		
				<%end if%>
			 </table>			
		   </div>	
		   
<!-- #include file="grid_bottom.asp" -->
</body>
</html>