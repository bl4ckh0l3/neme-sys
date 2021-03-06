<%
'************************************************************************************************
' SPECIFICARE LA LABEL DELLA CAMPAGNA VOUCHER DA UTILIZZARE PER GENERARE I CODICI VOUCHER
' DEVE ESSERE NECESSARIAMENTE UNA CAMPAIGN DI TIPO: "VOUCHER REGALO SINGOLO UTILIZZO"
' 
' SPECIFY THE LABEL OF THE VOUCHER CAMPAIGN TO BE USED TO GENERATE VOUCHER CODES 
' IT MUST NECESSARILY BE A CAMPAIGN OF TYPE "GIFT ONE SHOT"
'
voucher_campaign = ""
'
'
'************************************************************************************************



if (isEmpty(Session("objUtenteLogged"))) then
	response.Redirect(Application("baseroot")&"/login.asp?from="&Application("baseroot")&"/common/include/Controller.asp?gerarchia="&request("gerarchia"))
end if

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

' recupero url corrente per definire path a send.asp
tmpurl = request.ServerVariables("URL")
tmpurl = Mid(tmpurl,1,InStrRev(tmpurl,"/",-1,1))
      
if(Application("use_recaptcha") = 0) then
	'************* FUNZIONE PER IL VECCHIO CAPTCHA
  ' verifico che il codice captcha inserito dall'utente corrisponda con il captcha generato
  ' in caso contrario rimando alla pagine di registrazione con errore
  ' devo usare campo hidden perch� originale ( request("captchacode") ) non viene recuperato
  if not TestCaptcha("ASPCAPTCHA",  request("sent_captchacode")) then
    response.Redirect(tmpurl&"send.asp?captcha_err=1&gerarchia="&request("gerarchia")&"&id_news="&request("id_news"))
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
    response.Redirect(tmpurl&"send.asp?captcha_err=1&gerarchia="&request("gerarchia")&"&id_news="&request("id_news"))
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

Dim objUserLogged, objUserLoggedTmp
Set objUserLoggedTmp = new UserClass
Set objUserLogged = objUserLoggedTmp.findUserByID(Session("objUtenteLogged"))  
strEmail = objUserLogged.getEmail()
Set objUserLogged = nothing
Set objUserLoggedTmp = nothing

Dim sendmail, userMail, mailText, boolMailSent
Dim nome, cognome, telefono, indirizzo, zipcode, citta, nazione, infoBy
is_gift=1
friendMail = request("email")

subtmpurl = Mid(tmpurl,InStr(1,tmpurl,Application("dir_upload_templ"),1)+Len(Application("dir_upload_templ")),InStrRev(tmpurl,"/",-1,1))
templateContactPath = subtmpurl&"mail.asp"

errormessagetxt = ""


Set objVoucher = new VoucherClass
Set objDB = New DBManagerClass
Set objConn = objDB.openConnection()
objConn.BeginTrans
new_voucher_code = objVoucher.generateVoucherCodeByCampaignLabel(voucher_campaign, Session("objUtenteLogged"), objConn)
if objConn.Errors.Count = 0 then
  objConn.CommitTrans
else
  objConn.RollBackTrans
  response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
end if			
Set objDB = nothing
if(new_voucher_code<>"")then
  Set objMail = New SendMailClass
  call objMail.sendMailVoucher(new_voucher_code, is_gift, strEmail, templateContactPath, "", lang.getLangCode())
  Set objMail = Nothing
else
  errormessagetxt = lang.getTranslated("backend.voucher.label.error_generate_code") 
end if
Set objVoucher = nothing
%>
<!-- #include virtual="/common/include/setTemplateTargetList.inc" -->
<%
Dim bolHasObj
bolHasObj = false

on error Resume Next
if(id_news <> "")then
			Set objCurrentNews = News.findNewsByIDCached(id_news)
      bolHasObj = true  
elseif(id_news = "" AND bolCatContainNews AND not(isNull(objListaTargetCat))) then
  Set objListaNews = News.findNewsSlimCached(null, null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_by,false,false)	
  
  if(objListaNews.Count > 0) then
    Dim objSelNews, newsCounter, iIndex, objTmpNews, FromNews, ToNews, Diff
    iIndex = objListaNews.Count
    FromNews = ((numPage * elem_x_page) - elem_x_page)
    Diff = (iIndex - ((numPage * elem_x_page)-1))
    if(Diff < 1) then
      Diff = 1
    end if
    
    ToNews = iIndex - Diff
    
    totPages = iIndex\elem_x_page
    if(totPages < 1) then
      totPages = 1
    elseif((iIndex MOD elem_x_page <> 0) AND not ((totPages * elem_x_page) >= iIndex)) then
      totPages = totPages +1	
    end if		

    arrTmpListNews = objListaNews.Items
    Set objCurrentNews = arrTmpListNews(FromNews)    
  
    bolHasObj = true
  end if
end if    

  
if (Instr(1, typename(objCurrentNews), "NewsClass", 1) > 0) then
  if not(lang.getTranslated(objCurrentNews.getMetaDescription()) = "") then
    metaDescription = metaDescription & " " & lang.getTranslated(objCurrentNews.getMetaDescription())
  else
    metaDescription = metaDescription & " " & objCurrentNews.getMetaDescription()
  end if
  
  if not(lang.getTranslated(objCurrentNews.getMetaKeyword()) = "") then
    metaKeyword = metaKeyword & " " & lang.getTranslated(objCurrentNews.getMetaKeyword())
  else
    metaKeyword = metaKeyword & " " & objCurrentNews.getMetaKeyword()
  end if
  
  if not(lang.getTranslated(objCurrentNews.getPageTitle()) = "") then
    pageTemplateTitle = lang.getTranslated(objCurrentNews.getPageTitle())
  else
    if not(objCurrentNews.getPageTitle() = "")then
      pageTemplateTitle = objCurrentNews.getPageTitle()
    end if
  end if       
end if

metaDescription = Trim(metaDescription)
metaKeyword = Trim(metaKeyword)
pageTemplateTitle = Trim(pageTemplateTitle)

if (pageTemplateTitle = "") then
  pageTemplateTitle = lang.getTranslated("frontend.page.title")
end if      
    
if Err.number <> 0 then
  'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
  bolHasObj = false
end if	

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
