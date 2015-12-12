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
				Dim hasSmallImg
				hasSmallImg = false
				
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					
					select case objFileXNews.getFileTypeLabel()
					case 1
						hasSmallImg = true
						Exit for	
					case else
					end select
					Set objFileXNews = nothing	
				next
				
				if (cbool(hasSmallImg)) then%>
				<%' Lista immagini piccole
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)		
					if(objFileXNews.getFileTypeLabel() = 1) then
						if (objSelNews.getAbstract1() = "") then%>
							<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" hspace="0" vspace="0" border="0" align="top" alt="<%=objFileXNews.getFileDida()%>">
						<%elseif(Cint(objSelNews.getAbstract1()) > 0) then%>
							<img src="<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>" hspace="0" vspace="0" border="0" align="top" alt="<%=objFileXNews.getFileDida()%>">		
						<%elseif(Cint(objSelNews.getAbstract1()) = 0) then%>
							<div style="background-image:url(<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>);
							background-position:center top;background-repeat:no-repeat;width:77px;height:93px;text-align:center;
							color:#EE0B0B;font-size:12px;font-weight:bold;font-family:Arial Narrow;"><br/><br/><p valign="middle" style="background:#FFFFFF;"><%=lang.getTranslated("frontend.template_listenozze.label.prenotato")%></p></div>			
						<%end if
						objListaFilePerNews.remove(xObjFile)
						Exit for
					end if
					Set objFileXNews = nothing	
				next%>
				<%end if	
			end if
			Set objListaFilePerNews = nothing
		end if%>