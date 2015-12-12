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
				<table border="0" cellspacing="0" cellpadding="0" class="tablePhotoAttachment">
				<tr id="trImgPhoto">
				<%' Lista immagini grandi
				Dim imgBigCounter, bigImgID, xObjFileTmp
				imgBigCounter = 0
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)
					bigImgID = objFileXNews.getFileID()				
					if(objFileXNews.getFileTypeLabel() = 1) then
						for each xObjFileTmp in objListaFilePerNews
							Set objFileXNewsTmp = objListaFilePerNews(xObjFileTmp)
							if(objFileXNewsTmp.getFileTypeLabel() = 2) then
								if(strComp("zoom_"&objFileXNews.getFileName(), objFileXNewsTmp.getFileName(), 1) = 0) then
									bigImgID = objFileXNewsTmp.getFileID()
								end if
							end if
						next
						%>	
						<td id="img<%=objSelNews.getNewsID & imgBigCounter%>" class="tdPhotoAttachment" onmouseover="changeBackgroundImg('img<%=objSelNews.getNewsID & imgBigCounter%>',1,1);" onmouseout="changeBackgroundImg('img<%=objSelNews.getNewsID & imgBigCounter%>',0,0);"><a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popup.asp?id_allegato="&bigImgID%>','popupallegati',700,530,100,100)"><img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" width="101" height="67" hspace="0" vspace="0" border="0" alt="<%=objFileXNews.getFileDida()%>"></a></td>
						<%objListaFilePerNews.remove(xObjFile)
            			imgBigCounter = imgBigCounter +1
					end if
					Set objFileXNews = nothing	
				next%>
				</tr>
				</table><br><br>
				<%end if	
			end if
			Set objListaFilePerNews = nothing
		end if%>