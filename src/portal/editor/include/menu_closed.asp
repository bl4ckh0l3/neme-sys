<%
Dim menu_closed
menu_closed = 0

if(request("menu_closed")<>"") then
	menu_closed = request("menu_closed")
end if

Session("menu_closed") = menu_closed
%>