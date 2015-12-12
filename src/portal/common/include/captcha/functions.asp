<%
' The code below supplied by Mark Short 

' returns string the can be written where you would like the reCAPTCHA challenged placed on your page 
function recaptcha_challenge_writer(publickey) 
  recaptcha_challenge_writer = "<script type=""text/javascript"">" & _ 
  "var RecaptchaOptions = {" & _ 
  " theme : 'white'," & _ 
  " tabindex : 0" & _ 
  "};" & _ 
  "</script>" & _ 
  "<script type=""text/javascript"" src=""http://api.recaptcha.net/challenge?k=" & publickey & """></script>" & _ 
  "<noscript>" & _ 
  "<iframe src=""http://api.recaptcha.net/noscript?k=" & publickey & """ frameborder=""1""></iframe><br>" & _ 
  "<textarea name=""recaptcha_challenge_field"" rows=""3"" cols=""40""></textarea>" & _ 
  "<input type=""hidden"" name=""recaptcha_response_field"" value=""manual_challenge"">" & _ 
  "</noscript>" 
end function 

function recaptcha_confirm(privkey,rechallenge,reresponse) 
  ' Test the captcha field 
  Dim VarString 
  VarString = _ 
  "privatekey=" & privkey & _ 
  "&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & _ 
  "&challenge=" & rechallenge & _ 
  "&response=" & reresponse 
  Dim objXmlHttp 
  Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0") 
  objXmlHttp.open "POST", "http://api-verify.recaptcha.net/verify", False 
  objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
  objXmlHttp.send VarString 
  Dim ResponseString 
  ResponseString = split(objXmlHttp.responseText, vblf) 
  Set objXmlHttp = Nothing 
  if ResponseString(0) = "true" then 
    ' They answered correctly 
    recaptcha_confirm = "" 
  else 
    ' They answered incorrectly 
    recaptcha_confirm = ResponseString(1) 
  end if 
end function 
%>