<%
function TestCaptcha(byval valSession, byval valCaptcha)
	dim tmpSession
	valSession = Trim(valSession)
	valCaptcha = Trim(valCaptcha)
	if (valSession = vbNullString) or (valCaptcha = vbNullString) then
		TestCaptcha = false
	else
		tmpSession = valSession
		valSession = Trim(Session(valSession))
		Session(tmpSession) = vbNullString

		if valSession = vbNullString then
			TestCaptcha = false
		else
			valCaptcha = Replace(valCaptcha,"i","I")
			if StrComp(valSession,valCaptcha,1) = 0 then
				TestCaptcha = true
			else
				TestCaptcha = false
			end if
		end if		
	end if
end function


Dim id_news, News, objCurrentNews, strGerarchia, strCSS, objFileXNews, objListaFilePerNews
Dim objListaNews, objListaTargetCat, objListaTargetLang
Dim totPages, numPage, templateContactPath

      
if(Application("use_recaptcha") = 0) then
	'************* FUNZIONE PER IL VECCHIO CAPTCHA
  ' verifico che il codice captcha inserito dall'utente corrisponda con il captcha generato
  ' in caso contrario rimando alla pagine di registrazione con errore
  ' devo usare campo hidden perchè originale ( request("captchacode") ) non viene recuperato
  if not TestCaptcha("ASPCAPTCHA",  request("sent_captchacode")) then
    response.Redirect(Application("baseroot")&Application("dir_upload_templ")&"contactus/send.asp?captcha_err=1&gerarchia="&request("gerarchia")&"&id_news="&request("id_news"))
  end if
else
	'************* RECUPERO PARAMETRI RECAPTCHA
	Dim recaptcha_challenge_field, recaptcha_response_field, recaptcha_private_key, recaptcha_public_key, cTemp
	recaptcha_challenge_field  = request("sent_recaptcha_challenge_field")
	recaptcha_response_field   = request("sent_recaptcha_response_field")
	recaptcha_private_key      = Application("recaptcha_priv_key")
	recaptcha_public_key       = Application("recaptcha_pub_key")

	'************* CHECK VALORE RECAPTCHA
	cTemp = recaptcha_confirm(recaptcha_private_key, recaptcha_challenge_field, recaptcha_response_field)
	If cTemp <> "" Then 
    response.Redirect(Application("baseroot")&Application("dir_upload_templ")&"contactus/send.asp?captcha_err=1&gerarchia="&request("gerarchia")&"&id_news="&request("id_news"))
	end if
end if


Dim stato
stato = 1
if(request("is_preview_content")) then
	stato = null
end if

Set News = New NewsClass
id_news = request("id_news")
strGerarchia = request("gerarchia")
numPage = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

Dim sendmail, userMail, mailText, boolMailSent
Dim nome, cognome, telefono, indirizzo, zipcode, citta, nazione, infoBy
mailTo = request("mailTo")
userMail = request("email")
mailText =  request("testo")
nome = request("nome")
cognome = request("cognome")
telefono = request("telefono")
indirizzo = request("indirizzo")
zipcode = request("zipcode")
citta = request("citta")
nazione = request("nazione")

templateContactPath = "contactus/mail_contactus.asp"


Set objMail = New SendMailClass
call objMail.sendMailContactus(mailTo, userMail, mailText, nome, cognome, telefono, indirizzo, zipcode, citta, nazione, templateContactPath, lang.getLangCode())
Set objMail = Nothing
%>