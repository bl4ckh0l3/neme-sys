		<%if not(isNull(objCurrentNews.getFilePerNews())) AND not(isEmpty(objCurrentNews.getFilePerNews())) then
			Set objListaFilePerNews = objCurrentNews.getFilePerNews()
			
			if not(isEmpty(objListaFilePerNews)) then
				' LEGENDA TIPI FILE
				'2 = img big%>				
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
					end select
					Set objFileXNews = nothing	
				next%>
  
				<script language="JavaScript" type="text/javascript">
				function cambiaImg(cosa){
					document.getElementById('img').src = cosa
				}
				</script>
          
				<%' Lista immagini grandi
        Dim forCounter, strListAllegati
        forCounter = 0 
        strListAllegati = "" 
				for each xObjFile in objListaFilePerNews
					Set objFileXNews = objListaFilePerNews(xObjFile)					
					if(objFileXNews.getFileTypeLabel() = 2) then
              if(forCounter = 0) then  
                strListAllegati = strListAllegati & "<img src="""&Application("dir_upload_news")&objFileXNews.getFilePath()&""" alt="""&objFileXNews.getFileName()&""" name=""img"" border=""0"" id=""img"" hspace=""5"" vspace=""5""/><br /><a onclick=""cambiaImg('"&Application("dir_upload_news")&objFileXNews.getFilePath()&"')"" href=""#"">"&(forCounter+1)&"</a> - "
              else
                strListAllegati = strListAllegati & "<a onclick=""cambiaImg('"&Application("dir_upload_news")&objFileXNews.getFilePath()&"')"" href=""#"">"&(forCounter+1)&"</a> - "
              end if
            
            forCounter = forCounter + 1
            objListaFilePerNews.remove(xObjFile)
					end if
					Set objFileXNews = nothing	
				next
        response.write(Mid(strListAllegati,1,InStrRev(strListAllegati,"-",-1,1)-1))		
			end if
			Set objListaFilePerNews = nothing
		end if%>