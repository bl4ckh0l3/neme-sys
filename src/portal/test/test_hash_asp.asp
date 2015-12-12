<!-- #include virtual="/common/include/hex_sha1_js.asp" -->
<%
	Dim strPassWord, strHash
	strPassWord = "abc"
	strHash = hex_sha1(strPassWord)

	Response.Write("<p><b>strPassWord:</b> " & strPassWord & "</p>")
	Response.Write("<p><b>strHash:</b> " & strHash & "</p>")
%>