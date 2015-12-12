<%
'In IIS, 404 pages that are directed to an URL have the "error" URL attached in the query string.
'It looks something like this 404;http://www.me.com:80/code/nosuchfileblahblah1.asp
'We're gonna use it.. so grab it.
 
script = request.servervariables("QUERY_STRING")
if instr(script,"/") > 1 then
     myArray = split(script,"/")
       if instr(myArray(Ubound(myArray)),".asp") = 0 then
       myID = myArray(Ubound(myArray)) 'This is the method for obtaining the ID if you end your URL in the ID. Example: http://www.me.com/code/this-is-good-code/1
       else
       myID = replace(myArray(Ubound(myArray)),".asp","")  'This is the method for obtaining the ID if you end your URL in a fake extension. Example: http://www.me.com/code/this-is-good-code/1.asp
       end if
end if
 
'Now that you've extracted your code ID, just go about your business, you are done!
'Here's some sample code
 
if isNumeric(myID) then 'Make sure it's an ID and not some malicious code
response.write "The ID Extracted from the URL is: " & myID
else
response.redirect "http://www.me.com"
end if


'response.write Request.ServerVariables("SERVER_SOFTWARE")

%>