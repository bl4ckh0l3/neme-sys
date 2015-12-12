<% Option Explicit %>
<!-- #include file="inc.common.asp" -->
<%
	
	' 
	' $Id: frames.asp,v 1.1.1.1 2003/03/09 22:45:57 peter Exp $
	' 
	' 
	' 
	' @author	Peter Theill	peter@theill.com
	' 
	
	If (NOT loggedOn()) Then
		Response.Redirect "expired.asp?reload=true"
		Response.End
	End If
	
%>
<html>
<head>
	<title><%= getMsg("application.name") %></title>
	<link rel="stylesheet" type="text/css" href="css/chat.css">
	<script type="text/javascript" language="JavaScript1.2">
	<!--
		
		function showLogOffWindow() {
			executeRequest('action=logoff');
			
			/* 
			
			// you might enable this code if you will -force- a log out if the close
			// the browser window entirely without logging out. however most systems
			// are having popup blockers installed and this will disable the log out
			// anyway :-/
			
			var mConquerChatLogOut = window.open(
				"logout.asp",
				null,
				"toolbar=no,width=380,height=80,resizable=0"
			);
			
			mConquerChatLogOut.focus();
			*/
		}
		
		function onLoggedOff() {
			;
		}
		
	// -->
	</script>
</head>

<frameset rows="*,94" onUnload="showLogOffWindow()">
	<frameset cols="*,150">
		<frame name="messages" src="window.asp">
		<frameset rows="66%,34%,36">
			<frame name="users" src="users.asp" scrolling=no>
			<frame name="rooms" src="rooms.asp" scrolling=no>
			<frame name="add_room" src="addroom.asp" scrolling=no>
		</frameset>
	</frameset>
	
	<frame name="message" src="message.asp" scrolling=no noresize>
	
	<noframes>
	<body>
		ConquerChat is a frame-based free chat done in ASP and includes full source code.
	</body>
	</noframes>
</frameset>

</html>