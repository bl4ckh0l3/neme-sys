				<%
				'****************************** widget grafico preferenza utente
				'****************************** necessario parametro id_utente

				dim percentual, total, total_comment_news, total_comment_prod, objUserPreferenceP, objUsrLogTmpobjUsrLog, wIdUser
				Set objUsrLogTmp = new UserClass
				Set objUsrLog = objUsrLogTmp.findUserByID(Session("objUtenteLogged"))
				wIdUser=objUsrLog.getUserID()
        Set objUsrLog = nothing
        Set objUsrLogTmp = nothing

        percentual = 0
				Set objUserPreferenceP = new UserPreferenceClass
				percentual = objUserPreferenceP.findUserPreferencePositivePercent(wIdUser)
				percentual=FormatNumber(percentual, 0,-1)
				total = objUserPreferenceP.findNumUserPreferenceTotal(wIdUser, true)
        
				Set objComment = new CommentsClass
				total_comment_news = objComment.countCommentiByIDUtente(wIdUser,1,1)
				'<!--nsys-addsuser1-->
				total_comment_prod = objComment.countCommentiByIDUtente(wIdUser,2,1)
				'<!---nsys-addsuser1-->
				Set objComment = nothing
				%><br/>
				<div style="float:left;padding-left:5px;padding-top:10px;" class="txtUserPreference">
        <!--nsys-modcommunity7-->
				<%=lang.getTranslated("backend.utenti.detail.table.label.like")%>:&nbsp;<%=percentual%>%<br/>        
			<script type="text/javascript">
      $(function () {
        var chart;
        $(document).ready(function() {
          chart = new Highcharts.Chart({
            chart: {
              renderTo: 'usrprefchartbox',
              type: 'bar',
              width: 100,
              height: 70,
              spacingTop:-20,
              marginLeft:-1,
              marginRight:0
            },
            title: {
              text: ''
            },
            xAxis: {
              categories: [''],
              gridLineWidth:0
            },
            yAxis: {
              title: {
                text: ''
              },
              min: 0,
              max:100,
              showFirstLabel:false,
              showLastLabel:false,
              gridLineWidth:0
            },
            tooltip: {
              enabled: false
            },            
            legend: {
              enabled: false
            },
            series: [{
              data: [
              {
                color: 'blue',
                y: <%=percentual%>
              }
              ]
            }]
          });
        });
      });
			</script>
			<div align="left" id="usrprefchartbox" style="width:100px;height:5px;border:#000000 1px solid;overflow: hidden;"></div>
      <%
      endcounter=0      
      if(percentual>0 AND percentual<=20)then
        endcounter=1
      elseif(percentual>20 AND percentual<=40)then 
        endcounter=2
      elseif(percentual>40 AND percentual<=60)then
        endcounter=3
      elseif(percentual>60 AND percentual<=80)then 
        endcounter=4
      elseif(percentual>80 AND percentual<=100)then 
        endcounter=5
      end if
      if(endcounter>0)then
      %>
			<div align="left" id="usrprefstarsbox" style="width:100px;height:15px;">
      <%for starcount = 1 to endcounter%>
      <img width="14" height="15" src="<%=Application("baseroot")&"/common/img/ico_stella.png"%>" align="absmiddle" style="padding:0px;border:0px;">
      <%next%>
      </div><br/>
      <%end if%>
				<%=lang.getTranslated("backend.utenti.detail.table.label.total_vote")%>:&nbsp;<%=total%><br/>
      <!---nsys-modcommunity7-->
				<%=lang.getTranslated("backend.utenti.detail.table.label.total_commenti_news")%>:&nbsp;<%=total_comment_news%>
				<!--nsys-addsuser2--><br/><%=lang.getTranslated("backend.utenti.detail.table.label.total_commenti_prod")%>:&nbsp;<%=total_comment_prod%><!---nsys-addsuser2-->
				</div>
				<%
				Set objUserPreferenceP = nothing
				%>