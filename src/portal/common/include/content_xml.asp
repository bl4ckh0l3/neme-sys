<?xml version="1.0"?>
<%On Error Resume Next

Response.Charset="UTF-8"
Session.CodePage  = 65001
%>
<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<%
' Impostazione che setta il tipo di file in output su XML
response.ContentType = "text/xml"

Dim News, objListaNews, objListaTargetCat, objListaTargetLang
Dim strGerarchia, newsXpage, numPage, order_news_by

Set News = New NewsClass
strGerarchia = request("gerarchia")
order_news_by = 1

if(request("order_news_by")<>"")then
	order_news_by = request("order_news_by")
end if

newsXpage = 20
numPage = 1
stato = 1

'pageRssURL = ""
'Dim isHTTPS
'isHTTPS = Request.ServerVariables("HTTPS")
'If isHTTPS = "off" AND Application("use_https") = 1 Then
'	pageRssURL = "https://"&Request.ServerVariables("SERVER_NAME")& "/common/include/Controller.asp?gerarchia="&strGerarchia
'Else
'	pageRssURL = "http://"&Request.ServerVariables("SERVER_NAME")& "/common/include/Controller.asp?gerarchia="&strGerarchia
'End If
%>
<!-- #include virtual="/common/include/setTemplateTargetList.inc" -->
  <root> 
	<%
	'************** codice per la lista news e paginazione
	Dim bolHasObj
	bolHasObj = false
	
	on error Resume Next
	if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
		Set objListaNews = News.findNewsSlimCached(null, null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_news_by, false, true)	
		
		if(objListaNews.Count > 0) then		
			bolHasObj = true
		end if
	end if
		
	if Err.number <> 0 then
		bolHasObj = false
	end if			
	
	if(bolHasObj) then
		
		for each x in objListaNews
			Set objSelNews = objListaNews(x)%>
			<item>
			<title><![CDATA[<%=objSelNews.getTitolo()%>]]></title>
			<summary1><![CDATA[<%=objSelNews.getAbstract1()%>]]></summary1>
			<summary2><![CDATA[<%=objSelNews.getAbstract2()%>]]></summary2>
			<summary3><![CDATA[<%=objSelNews.getAbstract3()%>]]></summary3>
			<description><![CDATA[<%=objSelNews.getTesto()%>]]></description>
			<keyword><![CDATA[<%=objSelNews.getKeyword()%>]]></keyword>
			<page_title><![CDATA[<%=objSelNews.getPageTitle()%>]]></page_title>
			<meta_keyword><![CDATA[<%=objSelNews.getMetaKeyword()%>]]></meta_keyword>
			<meta_description><![CDATA[<%=objSelNews.getMetaDescription()%>]]></meta_description>
			<data_public><![CDATA[<%=objSelNews.getDataPubNews()%>]]></data_public>
			<%
			if not(isNull(objSelNews.getFilePerNews())) AND not(isEmpty(objSelNews.getFilePerNews())) then
				Set objListaFilePerNews = objSelNews.getFilePerNews()

				if not(isEmpty(objListaFilePerNews)) then
					' LEGENDA TIPI FILE      
					'1 = img small
					'2 = img big
					'3 = pdf
					'4 = audio-video
					'5 = others...
					'6 = img medium
					'7 = img carrello   

					' Lista label tipi file
					Dim hasSmallImg, hasMediumImg, hasBigImg, hasCardImg, hasPdf, hasAudioVideo, hasOthers
					hasSmallImg = false
					hasBigImg = false
					hasPdf = false
					hasAudioVideo = false
					hasOthers = false
					hasMediumImg = false
					hasCardImg = false

					bolHasAttach = true

					Set attachMap = Server.CreateObject("Scripting.Dictionary")
					Set attachMultiLangKey = Server.CreateObject("Scripting.Dictionary")
					Set attachSmall = Server.CreateObject("Scripting.Dictionary")
					Set attachBig = Server.CreateObject("Scripting.Dictionary")
					Set attachPdf = Server.CreateObject("Scripting.Dictionary")
					Set attachAudioVideo = Server.CreateObject("Scripting.Dictionary")
					Set attachOther = Server.CreateObject("Scripting.Dictionary")
					Set attachMedium = Server.CreateObject("Scripting.Dictionary")
					Set attachCard = Server.CreateObject("Scripting.Dictionary")

					for each xObjFile in objListaFilePerNews
						Set objFileXNews = objListaFilePerNews(xObjFile)					

						select case objFileXNews.getFileTypeLabel()
						case 1
							hasSmallImg = true
							attachSmall.add objFileXNews, ""
						case 2
							hasBigImg = true
							attachBig.add objFileXNews, ""	
						case 3
							hasPdf = true
							attachPdf.add objFileXNews, ""
						case 4
							hasAudioVideo = true
							attachAudioVideo.add objFileXNews, ""
						case 5
							hasOthers = true
							attachOther.add objFileXNews, ""
						case 6
							hasMediumImg = true
							attachMedium.add objFileXNews, ""
						case 7						
							hasCardImg = true
							attachCard.add objFileXNews, ""
						case else
						end select
						
						Set objFileXNews = nothing	
					next

					attachMap.add "small", attachSmall
					attachMap.add "big", attachBig
					attachMap.add "pdf", attachPdf
					attachMap.add "media", attachAudioVideo
					attachMap.add "other", attachOther
					attachMap.add "medium", attachMedium
					attachMap.add "card", attachCard

					attachMultiLangKey.add "small", "frontend.file_allegati.label.key_img_small"
					attachMultiLangKey.add "big", "frontend.file_allegati.label.key_img_big"
					attachMultiLangKey.add "pdf", "frontend.file_allegati.label.key_pdf"
					attachMultiLangKey.add "media", "frontend.file_allegati.label.key_audio_video"
					attachMultiLangKey.add "other", "frontend.file_allegati.label.key_others_doc"
					attachMultiLangKey.add "medium", "frontend.file_allegati.label.key_img_medium"
					attachMultiLangKey.add "card", "frontend.file_allegati.label.key_img_card"        
				end if
				
				Set objListaFilePerNews = nothing
			end if 
			
			if(bolHasAttach) then%> 
				<attachments>
				<%for each key in attachMap
					if(attachMap(key).count > 0)then%>
						<%for each item in attachMap(key)%>
							<attach><%=Application("dir_upload_news")&item.getFilePath()%></attach>
						<%next
					end if
				next%>
				</attachments>
			<%end if%>			
			</item>
			<%Set objSelNews = nothing
		next		
	end if%>
  </root> 

<%
Set objListaTargetCat = nothing
Set objListaTargetLang = nothing
Set objListaNews = nothing
Set News = Nothing

if(Err.number <> 0) then
	response.write(Err.description)
end if
%> 