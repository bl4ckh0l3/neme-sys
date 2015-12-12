		<%if not(isNull(objSelNews.getFilePerNews())) AND not(isEmpty(objSelNews.getFilePerNews())) then
			Set objListaFilePerNews = objSelNews.getFilePerNews()
			
			if not(isEmpty(objListaFilePerNews)) then
				' LEGENDA TIPI FILE
				'1 = img small
				'2 = img big
				'3 = pdf
				'4 = audio-video
				'5 = others%>				
				<%
				' Lista label tipi file
				Dim hasBigImg
				hasBigImg = false
				
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					
					select case objFileXNews.getFileTypeLabel()
					case 2
						hasBigImg = true
						Exit for	
					case else
					end select
					Set objFileXNews = nothing	
				next
				
				if (cbool(hasBigImg)) then%>
				<%' Lista immagini grandi
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)		
					if(objFileXNews.getFileTypeLabel() = 2) then%><img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" hspace="0" vspace="0" border="0" align="top" alt="<%=objFileXNews.getFileDida()%>"><%objListaFilePerNews.remove(xObjFile)
					end if
					Set objFileXNews = nothing	
				next%><%end if	
			end if
			Set objListaFilePerNews = nothing
		end if%>