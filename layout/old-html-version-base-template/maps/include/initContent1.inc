		<%
		'************** codice per la lista news e paginazione
		Dim bolHasObj
		bolHasObj = false
		
		on error Resume Next
		if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
			Set objListaNews = News.findNews(null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_news_by)	
			
			if(objListaNews.Count > 0) then		
				bolHasObj = true
			end if
		end if
			
		if Err.number <> 0 then
			bolHasObj = false
		end if			
		
		if(bolHasObj) then%>
        <script type="text/javascript">
        <!--
    
        function selectDetail(value, index){

    
          var el = document.getElementById("maps-list-selected"+value); 
    
          if(el.style.visibility== "visible"){
            el.style.visibility= "hidden";
            el.style.display= "none";
          }else{
            el.style.visibility= "visible"; 
            el.style.display= "block";
          } 

          for(var i=0; i<=index;i++)
          {
            if("maps-list-selected"+i != "maps-list-selected"+value){
              document.getElementById("maps-list-selected"+i).style.visibility = "hidden";
              document.getElementById("maps-list-selected"+i).style.display = "none";
            }
          }
          
          //document.getElementById("maps-list-selected"+value).style.visibility = "visible";
          //document.getElementById("maps-list-selected"+value).style.display = "block";      
        }
        //-->
        </script> 
        <br/>
        <div id="maps-list">
          <ul>
          <%Dim objSelNews, newsCounter, iIndex, objTmpNews              
          objTmpNews = objListaNews.Items
          
          iIndex = objListaNews.Count-1
          
          for newsCounter = 0 to iIndex
            Set objSelNews = objTmpNews(newsCounter)%>               
              <li><a href="#" onclick='javascript:selectDetail("<%=newsCounter%>",<%=iIndex%>);'><%=objSelNews.getTitolo()%></a></li>
              <%if ((newsCounter+1) MOD 4 = 0) then%>                
                </ul><ul>        
              <%end if%>
            <%Set objSelNews = nothing
          next%>
          </ul>  
			  </div> 
  
          <%for newsCounter = 0 to iIndex
            Set objSelNews = objTmpNews(newsCounter)%> 
              <div id="maps-list-selected<%=newsCounter%>" style="visibility:hidden;display:none;background-color:#cccccc;padding:5px 5px 5px 5px;">
                <h2 style="font-weight:bold;"><%=objSelNews.getTitolo()%></h2>
                <%=objSelNews.getTesto()%>
              </div>
            <%Set objSelNews = nothing
          next%>
              
		<%else
			response.Write("<br/><br/><div align=""center""><strong>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</strong></div>")
		end if%>