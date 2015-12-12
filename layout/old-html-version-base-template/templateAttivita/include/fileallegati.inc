		<%if not(isNull(objCurrentNews.getFilePerNews())) AND not(isEmpty(objCurrentNews.getFilePerNews())) then
			Set objListaFilePerNews = objCurrentNews.getFilePerNews()
			
			if not(isEmpty(objListaFilePerNews)) then
				' LEGENDA TIPI FILE
				'1 = img small
				'2 = img big
				'3 = pdf
				'4 = audio-video
				'5 = others			
				
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
				
          Dim counter
				if (cbool(hasBigImg)) then
          ' Lista immagini grandi
          counter = 0%>              
            <img src="<%=Application("baseroot")&"/common/img/spacer.gif"%>"  hspace="0" vspace="0" border="0" name="SlideShow" id="SlideShow">           
            <script>        
            <%for each xObjFile in objListaFilePerNews
              Set objFileXNews = objListaFilePerNews(xObjFile)					
              if(objFileXNews.getFileTypeLabel() = 2) then%>
 
                  pictures[<%=counter%>] = '<%=Application("dir_upload_news")&objFileXNews.getFilePath()%>';           
                <%objListaFilePerNews.remove(xObjFile)
              end if
              Set objFileXNews = nothing	
              counter = counter + 1
            next%>        
              

              var preLoad = new Array()
              for (i = 0; i < pictures.length; i++){
                 preLoad[i] = new Image()
                 preLoad[i].src = pictures[i];              
              }   
             
            var arrSlideshowFunction = new Array("runSlideShow();","StartAnim();","slideSwhowFading();","runSlideShow();", "StartAnim();", "slideSwhowFading();"); 
            var functChooiceNum = Math.floor(Math.random()*5+0);
            var chooice = arrSlideshowFunction[functChooiceNum];       

              //alert("functChooiceNum: " + functChooiceNum+ " chooice: " + chooice);
              
            eval(chooice);  
            </script>           
				<%end if

			end if
			Set objListaFilePerNews = nothing
		end if%>