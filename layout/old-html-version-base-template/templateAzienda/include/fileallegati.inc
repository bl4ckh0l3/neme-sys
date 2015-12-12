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
					case 2
						hasBigImg = true	
					case 3
						hasPdf = true
					case 4
						hasAudioVideo = true
					case 5
						hasOthers = true
					case else
					end select
					Set objFileXNews = nothing	
				next
				
				' Lista immagini grandi
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					if(objFileXNews.getFileTypeLabel() = 2) then%>
						<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" hspace="0" vspace="6" border="0" alt="<%=objFileXNews.getFileDida()%>"><br/>
						<%objListaFilePerNews.remove(xObjFile)
					end if
					Set objFileXNews = nothing	
				next
				
				if (cbool(hasAudioVideo)) then response.write("<br/><b>"&lang.getTranslated("frontend.file_allegati.label.key_audio_video")&"</b><br/><br/>") end if
				' Lista audiovideo
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					if(objFileXNews.getFileTypeLabel() = 4) then%>
						<a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popup.asp?id_allegato="&objFileXNews.getFileID()%>','popupallegati',400,400,100,100)"><%=objFileXNews.getFileName()%></a><br><br>					
						<%objListaFilePerNews.remove(xObjFile)
					end if
					Set objFileXNews = nothing	
				next
				
				' Lista pdf
				if (cbool(hasPdf)) then%><img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>" height="58" vspace="0" hspace="0" border="0" align="top"><br><%end if
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					if(objFileXNews.getFileTypeLabel() = 3) then%>
						<img src="<%=Application("baseroot")&"/common/img/pulsante_download.gif"%>" hspace="0" vspace="5" border="0" align="center"><br/>
						<a target="_blank" href="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" class="downloadPdfDida"><%=objFileXNews.getFileDida()%></a><br>					
						<%objListaFilePerNews.remove(xObjFile)
					end if
					Set objFileXNews = nothing	
				next
				if (cbool(hasPdf)) then%><a target="_blank" href="http://www.adobe.com/it/products/acrobat/readstep2.html"><img src="<%=Application("baseroot")&"/common/img/pulsante_acrobat.gif"%>" vspace="10" hspace="0" border="0" align="top"></a><br><%end if
		
				
				if (cbool(hasOthers)) then response.write("<br/><b>"&lang.getTranslated("frontend.file_allegati.label.key_others_doc")&"</b><br/><br/>") end if
				' Lista others documents
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					if(objFileXNews.getFileTypeLabel() = 5) then%>
						<a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popup.asp?id_allegato="&objFileXNews.getFileID()%>','popupallegati',400,400,100,100)"><%=objFileXNews.getFileName()%></a><br><br>					
						<%objListaFilePerNews.remove(xObjFile)
					end if
					Set objFileXNews = nothing	
				next				
			end if
			Set objListaFilePerNews = nothing
		end if%>