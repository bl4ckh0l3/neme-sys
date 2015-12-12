		<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr>
		<td><img src="<%=Application("baseroot")&"/templates/templateShowroom/img/base_rollover_big.jpg"%>" vspace="6" hspace="9" border="0" id="showroomRollover" name="showroomRollover"></td>
		</tr>
		<tr>
		<td align="center">
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
				Dim hasSmallImg, hasBigImg
				hasSmallImg = false
				hasBigImg = false
				
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					
					select case objFileXNews.getFileTypeLabel()
					case 1
						hasSmallImg = true
					case 2
						hasBigImg = true	
					case else
					end select
					Set objFileXNews = nothing	
				next
				
				if (cbool(hasSmallImg)) then%>
				<script type="text/javascript" language="JavaScript">
				<!--
				<%' Lista immagini grandi
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)
					if(objFileXNews.getFileTypeLabel() = 2) then%>	
					Buffer("<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>");
					<%end if
					Set objFileXNews = nothing	
				next%>			
				//-->
				</script>
				<%' Lista immagini piccole
				Dim imgBigCounter, bigImgID, xObjFileTmp
				imgBigCounter = 0
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)
					bigImgID = objFileXNews.getFileID()				
					Dim srcFileBig
					if(objFileXNews.getFileTypeLabel() = 1) then						
						srcFileBig = Replace(objFileXNews.getFilePath(),".jpg","_zoom.jpg")%>	
						<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" width="70" height="47" hspace="0" vspace="1" border="0" alt="<%=objFileXNews.getFileDida()%>" onmouseover="changeBackgroundImg('showroomRollover','<%=Application("dir_upload_news")&srcFileBig%>');" onmouseout="changeBackgroundImg('showroomRollover','<%=Application("baseroot")&"/templates/templateShowroom/img/base_rollover_big.jpg"%>');">
						<%
						if(imgBigCounter = 6) then response.write("<br/>")
						%>
						<%objListaFilePerNews.remove(xObjFile)
            			imgBigCounter = imgBigCounter +1
					end if
					Set objFileXNews = nothing	
				next%>
				<%end if	
			end if
			Set objListaFilePerNews = nothing
		end if%>
		</td></tr></table>