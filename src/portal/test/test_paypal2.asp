<%
Response.ContentType = "text/xml"
%>
<BkItem>
<BookID>4356</BookID>
<Title>il signore degli anelli</Title>
<Writer>tolkien</Writer>
<Stock>fantascienza</Stock>
<Price><%=request("amount")%></Price>
</BkItem>