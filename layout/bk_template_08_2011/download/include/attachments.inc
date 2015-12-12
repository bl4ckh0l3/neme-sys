		<%if not(isNull(objCurrentNews.getFilePerNews())) AND not(isEmpty(objCurrentNews.getFilePerNews())) then
			Set objListaFilePerNews = objCurrentNews.getFilePerNews()
			
			if not(isEmpty(objListaFilePerNews)) then
				' LEGENDA TIPI FILE
				'1 = img small
				'2 = img big
				'3 = pdf
				'4 = audio-video
				'5 = others%>				
				<%
				' Lista label tipi file
				Dim hasBigImg, hasPdf, hasAudioVideo, hasOthers
				hasBigImg = false
				hasPdf = false
				hasAudioVideo = false
				hasOthers = false
				
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					
					select case objFileXNews.getFileTypeLabel()
					case 3
						hasPdf = true
					case 5
						hasOthers = true
					case else
					end select
					Set objFileXNews = nothing	
				next
				
				if (cbool(hasPdf)) then response.write("<br/><strong>"&lang.getTranslated("frontend.file_allegati.label.key_pdf")&"</strong><br/><br/>") end if
				' Lista pdf
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					if(objFileXNews.getFileTypeLabel() = 3) then%>
						<a class="pdfAttachLink" href="javascript:openWin('<%=Application("baseroot")&Application("dir_upload_templ")&"download/popup_download.asp?id_allegato="&objFileXNews.getFileID()%>','popupallegati',400,400,100,100)"><%=objFileXNews.getFileName()%></a><br><br>					
						<%objListaFilePerNews.remove(xObjFile)
					end if
					Set objFileXNews = nothing	
				next
				
				if (cbool(hasOthers)) then response.write("<br/><strong>"&lang.getTranslated("frontend.file_allegati.label.key_applications")&"</strong><br/><br/>") end if
				' Lista others documents
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					if(objFileXNews.getFileTypeLabel() = 5) then%>
						<a href="javascript:openWin('<%=Application("baseroot")&Application("dir_upload_templ")&"download/popup_download.asp?id_allegato="&objFileXNews.getFileID()&"&force=1"%>','popupallegati',400,400,100,100)"><%=objFileXNews.getFileName()%></a><br><br>					
						<%objListaFilePerNews.remove(xObjFile)
					end if
					Set objFileXNews = nothing	
				next				
			end if
			Set objListaFilePerNews = nothing
		end if%>