<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=lang.getTranslated("frontend.page.title")%></title>
<meta name="autore" content="Neme-sys; email:info@neme-sys.org">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/area_user.css"%>" type="text/css">
<script type="text/JavaScript" src="<%=Application("baseroot") & "/common/js/swfobject.js"%>"></script> 
<script language="JavaScript">

function changeTab(number){
	if(number==1)
		location.href='<%=Application("baseroot") & "/area_user/userprofile.asp"%>';
	else if(number==2)
		location.href='<%=Application("baseroot") & "/area_user/manageuser.asp"%>';
	else if(number==3)
		location.href='<%=Application("baseroot") & "/area_user/friendlist.asp"%>';
	else if(number==4)
		location.href='<%=Application("baseroot") & "/area_user/userphotos.asp"%>';

}


      
function checkAjaxHasFriendActiveP(divprofile, divname, id_friend, usrnameCurrUser){
  var query_string = "id_utente="+id_friend+"&action=2";

  $.ajax({
     type: "POST",
     cache: false,
     url: "<%=Application("baseroot") & "/area_user/checkfriend.asp"%>",
     data: query_string,
      success: function(response) {
        // show friend request icon
        //alert("response: "+response);
        if(response!=1){
				$("#"+divprofile+id_friend).hide();
				$("#"+divname+id_friend).empty().append(usrnameCurrUser);
				}else{
				$("#"+divname+id_friend).empty();
				$("#"+divprofile+id_friend).show();					
				}
      },
      error: function() {
				$(""+divprofile+id_friend).hide();
				$("#"+divname+id_friend).empty().append(usrnameCurrUser);
      }
   });
}

$(function() {
	$("#send-message").draggable();
});

var tempX = 0;
var tempY = 0;

jQuery(document).ready(function(){
	$(document).mousemove(function(e){
	tempX = e.pageX;
	tempY = e.pageY;
	}); 
})
		
function showDiv(elem){
	var divelement = document.getElementById(elem);
	var jquery_id= "#"+elem;
	hideAllDiv();
	$(jquery_id).show(1000);
	divelement.style.visibility = "visible";
	divelement.style.display = "block";
}
		
function hideAllDiv(){
	var divelement1 = document.getElementById("bacheca");
	var divelement2 = document.getElementById("statistiche");
	var divelement3 = document.getElementById("info");
	var divelement4 = document.getElementById("commenti");
	var divelement5 = document.getElementById("foto");
	divelement1.style.visibility = "hidden";
	divelement1.style.display = "none";
	divelement2.style.visibility = "hidden";
	divelement2.style.display = "none";
	divelement3.style.visibility = "hidden";
	divelement3.style.display = "none";
	divelement4.style.visibility = "hidden";
	divelement4.style.display = "none";
	divelement5.style.visibility = "hidden";
	divelement5.style.display = "none";
}

function prepareMessage(){		
	var divmessage = document.getElementById("send-message");
      
      if(ie||mac_ie){
        divmessage.style.left=tempX+10;
        divmessage.style.top=tempY+10;
      }else{
        divmessage.style.left=tempX+10+"px";
        divmessage.style.top=tempY+10+"px";
      }

	$("#send-message").show(1000);
	divmessage.style.visibility = "visible";
	divmessage.style.display = "block";
}
		
function insertMessage(){
	document.form_message.submit();      
}
		
function hideMessageform(){
	var divmessage = document.getElementById("send-message");
	divmessage.style.visibility = "hidden";
	divmessage.style.display = "none";
}
</script>
</head>
<body>
<!-- #include virtual="/fckeditor/fckeditor.asp" -->	
<%
'*************** INIZIALIZZO IL CODICE PER GENERARE GLI EDITOR HTML
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.Width = 350
oFCKeditor.Height = 200
oFCKeditor.ToolbarSet ="Simple"
oFCKeditor.BasePath = "/fckeditor/"
%>
<!-- #include file="grid_top.asp" -->

		<script language="JavaScript">

		</script>

        <h1><%=lang.getTranslated("frontend.header.label.utente_profile")%>&nbsp;<em><%=strUserName%></em></h1>

    	<p style="padding-bottom:10px;">
		<a href="javascript:showDiv('bacheca');"><b><%=lang.getTranslated("frontend.area_user.label.bacheca")%></b></a>&nbsp;|&nbsp;
		<a href="javascript:showDiv('commenti');"><b><%=lang.getTranslated("frontend.area_user.label.commenti")%></b></a>&nbsp;|&nbsp;
		<a href="javascript:showDiv('statistiche');"><b><%=lang.getTranslated("frontend.area_user.label.statistiche")%></b></a>&nbsp;|&nbsp;
		<a href="javascript:showDiv('info');"><b><%=lang.getTranslated("frontend.area_user.label.info_presonali")%></b></a>&nbsp;|&nbsp;
		<a href="javascript:showDiv('foto');"><b>foto<%=lang.getTranslated("frontend.area_user.label.foto")%></b></a>
		</p>
		<%
		'*********** inizializzo oggetti utili
		Set objUserPreference = new UserPreferenceClass	
    Set objComment = new CommentsClass
		Set objUserClass = new UserClass		
		%>

		<!--***************************************** TAB BACHECA *****************************************-->

		<div id="bacheca">   
		
      &nbsp;&nbsp;<a href="javascript:prepareMessage();"><img id="ok" alt="<%=lang.getTranslated("portal.templates.commons.label.vote_up")%>" src="<%=Application("baseroot") & "/common/img/comment_add.png"%>"/></a>
			<%      
			Dim id_utente_message,message_answ, vote_answ
			id_utente_message = request("id_utente_message")
			message_answ = request("message")
			vote_answ = request("vote")

				if(request("message_sent") = "1") then
        
          Set objDB = New DBManagerClass
          Set objConn = objDB.openConnection()
          objConn.BeginTrans
          call objUserPreference.insertUserPreference(id_utente_message, Session("objUtenteLogged"), null, null, vote_answ, message_answ, objConn)
                    
          if objConn.Errors.Count = 0 then
            objConn.CommitTrans%>
            <span id="vote-confirmed"><%=lang.getTranslated("portal.templates.commons.label.message_done")%></span><br/>
          <%else
            objConn.RollBackTrans%>
            <span id="vote-confirmed"><%=lang.getTranslated("portal.templates.commons.label.message_not_done")%></span><br/>
          <%end if			
          Set objDB = nothing
        end if
			%>
			<div id="send-message" style="left:0px;top:0px;z-index:1000;position:absolute;margin-bottom:3px;vertical-align:top;text-align:left;font-size: 10px;text-decoration: none;visibility:hidden;display:none;border:1px solid;padding:10px;background:#FFFFFF;width:350px;">
				<form action="<%=Application("baseroot") & "/area_user/publicprofile.asp"%>" method="post" name="form_message" accept-charset="UTF-8">		
				<input type="hidden" name="message_sent" value="1">  
				<input type="hidden" name="id_utente_message" value="<%=request("id_utente")%>"> 
				<input type="hidden" name="id_utente" value="<%=request("id_utente")%>">
				<p align="right"><a href="javascript:hideMessageform();">x</a></p>
        
				<div style="float:top;"><span class="labelForm"><%=lang.getTranslated("portal.templates.commons.label.insert_vote")%></span><br>
				<!--<textarea name="message" class="formFieldTXTTextareaComment" ></textarea>-->
				<%
				oFCKeditor.Value = ""
				oFCKeditor.Create "message"
				%></div>
				<%if not(id_utente = Session("objUtenteLogged"))then%>
				<div><span><%=lang.getTranslated("frontend.area_user.manage.label.like")%></span><br>
				<select name="vote" id="vote">
					<OPTION VALUE="-1"></OPTION>
					<OPTION VALUE="1" <%if (strComp("1", bolPublic, 1) = 0) then response.Write("selected")%>><%=lang.getTranslated("portal.commons.yes")%></OPTION>
					<OPTION VALUE="0" <%if (strComp("0", bolPublic, 1) = 0) then response.Write("selected")%>><%=lang.getTranslated("portal.commons.no")%></OPTION>
				</select>&nbsp;&nbsp;
				<input name="send" style="margin-left:70px;" value="<%=lang.getTranslated("frontend.area_user.manage.label.insert_message")%>" type="button" onclick="javascript:insertMessage();">		
				</div>
				<%else%>
					<input type="hidden" name="vote" value="-1">  
					<input name="send" value="<%=lang.getTranslated("frontend.area_user.manage.label.insert_message")%>" type="button" onclick="javascript:insertMessage();">
				<%end if%>
          
				</form>
			</div>

		<%On Error Resume Next    
		Dim idTmpCommento, idTmpCommentoOld, preferenceMessagesList, typeTmpPreference, intCountPref
		Set preferenceMessagesList = objUserPreference.getListUserPreferenceByUserFiltered(id_utente,true,true)
		idTmpCommentoOld = -1
		intCountPref = 0
		
		if (Instr(1, typename(preferenceMessagesList), "Dictionary", 1) > 0) then
			for each j in preferenceMessagesList
			  Set objPreference = preferenceMessagesList(j)
			  typeTmpPreference = objPreference.getTypeCommento()
		
			  Set objUserComment = objUserClass.findUserByID(objPreference.getIdFriend())%>
				<div style="padding-left:30px;padding-bottom:5px;">
				<span style="font-size:10px;"><%=objPreference.getInsertDate()&" - "%></span>
				<strong>
				<%if(objUserComment.getPublic()) then%>
				  <span id="showprofilep<%=intCountPref%>_<%=objUserComment.getUserID()%>"><a title="<%=lang.getTranslated("portal.templates.commons.label.view_pub_profile")%>" href="<%=Application("baseroot") & "/area_user/publicprofile.asp?id_utente="&objUserComment.getUserID()%>"><%=objUserComment.getUsername()%></a></span>
				  <span id="shownamep<%=intCountPref%>_<%=objUserComment.getUserID()%>"></span>
				  <script>
						$("#showprofilep<%=intCountPref%>_<%=objUserComment.getUserID()%>").hide();                  
				  checkAjaxHasFriendActiveP('showprofilep<%=intCountPref%>_','shownamep<%=intCountPref%>_',<%=objUserComment.getUserID()%>, '<%=objUserComment.getUsername()%>');
				  </script>
				<%else%>
				  <%=objUserComment.getUsername()%>
				<%end if%>
				</strong>				
				<br><%if(objPreference.getType()=1)then%><img id="nolike" src="<%=Application("baseroot") & "/common/img/like.png"%>" align="absbottom"/><%else if(objPreference.getType()=0)then%><img id="nolike" src="<%=Application("baseroot") & "/common/img/nolike.png"%>" align="absbottom"/><%end if end if%>
				<%=objPreference.getValue()%>
				</div>
			  <%Set objPreference = nothing
			  
			  idTmpCommentoOld = idTmpCommento
			  Set objUserComment = nothing
			  intCountPref = intCountPref+1
			next
			
			Set preferenceMessagesList = nothing
		end if
		
		if(Err.number <> 0)then
		'response.write(Err.description)
		end if
		%>
		</div>

		<!--***************************************** TAB COMMENTI *****************************************-->
	
		<div id="commenti" style="visibility:hidden;display:none;">
		<%On Error Resume Next    
		Dim idTmpCommento2, idTmpCommentoOld2, preferenceMessagesList2, typeTmpPreference2
		Set preferenceMessagesList2 = objUserPreference.getListUserPreferenceByUserFiltered(id_utente,false,false)
		idTmpCommentoOld2 = -1
		intCountPref = 0
		
		if (Instr(1, typename(preferenceMessagesList2), "Dictionary", 1) > 0) then
			for each j in preferenceMessagesList2
				Set objPreference2 = preferenceMessagesList2(j)
				idTmpCommento2 = objPreference2.getIdCommentoUser()
				typeTmpPreference2 = objPreference2.getTypeCommento()
				
				Set objUserComment = objUserClass.findUserByID(objPreference2.getIdFriend())
				if(idTmpCommento2 <> idTmpCommentoOld2) then
					Set objSelectedCommento = objComment.findCommentiByIDCommento(idTmpCommento2,typeTmpPreference2,1)
					response.write("<span style='font-size:10px;'>"&objSelectedCommento.getDtaInserimento()&"</span>&nbsp;-&nbsp;<strong>"&objUserClass.findUserByID(objSelectedCommento.getIDUtente()).getUsername()&"</strong><br/>" & objSelectedCommento.getMessage())
					Set objSelectedCommento = nothing
				end if%>
				<div style="padding-left:30px;padding-bottom:5px;"><span style="font-size:10px;"><%=objPreference2.getInsertDate()&"</span> - "%>
				<strong>
		    <%if(objUserComment.getPublic()) then%>
		      <span id="showprofilec<%=intCountPref%>_<%=objUserComment.getUserID()%>"><a title="<%=lang.getTranslated("portal.templates.commons.label.view_pub_profile")%>" href="<%=Application("baseroot") & "/area_user/publicprofile.asp?id_utente="&objUserComment.getUserID()%>"><%=objUserComment.getUsername()%></a></span>
		      <span id="shownamec<%=intCountPref%>_<%=objUserComment.getUserID()%>"></span>
		      <script>
					$("#showprofilec<%=intCountPref%>_<%=objUserComment.getUserID()%>").hide();                  
		      checkAjaxHasFriendActiveP('showprofilec<%=intCountPref%>_','shownamec<%=intCountPref%>_',<%=objUserComment.getUserID()%>, '<%=objUserComment.getUsername()%>');
		      </script>
		    <%else%>
		      <%=objUserComment.getUsername()%>
		    <%end if%>
				</strong>
				<br>
				<%if(objPreference2.getType()=1)then%><img id="like2" src="<%=Application("baseroot") & "/common/img/like.png"%>" align="absbottom"/><%else if(objPreference2.getType()=0)then%><img id="nolike2" src="<%=Application("baseroot") & "/common/img/nolike.png"%>" align="absbottom"/><%end if end if%>
				<%=objPreference2.getValue()%></div>

				<% Set objPreference2 = nothing
				
				idTmpCommentoOld2 = idTmpCommento2
				Set objUserComment = nothing
			  intCountPref = intCountPref+1
			next
			
			Set preferenceMessagesList2 = nothing
		end if
		
		if(Err.number <> 0)then
		'response.write(Err.description)
		end if
		%>
		</div>

		<!--***************************************** TAB STATISTICHE *****************************************-->

  		<div id="statistiche" style="visibility:hidden;display:none;">
		<%
		' widget grafico preferenza utente
		dim percentual_u, total_u, total_comment_news_u, total_comment_prod_u
		Dim total_news_u, total_prod_u
		percentual_u = 0
		percentual_u = objUserPreference.findUserPreferencePositivePercent(id_utente)
		percentual_u = FormatNumber(percentual_u, 0,-1)
		total_u = objUserPreference.findNumUserPreferenceTotal(id_utente, true)

		total_comment_news_u = objComment.countDistinctCommentiByIDUtente(id_utente,1,1)    
		Set obiNews = new NewsClass
		total_news_u = obiNews.countNews(null, null, null, null, null, null, null, null, null)
		total_news_u = FormatNumber(Cint(total_comment_news_u)*100/Cint(total_news_u), 0,-1)			
		Set obiNews = nothing

		'<!--nsys-publicprofile1-->
		total_comment_prod_u = objComment.countDistinctCommentiByIDUtente(id_utente,2,1)
		Set objProd = new ProductsClass
		total_prod_u = objProd.countProdotti()
		total_prod_u = FormatNumber(Cint(total_comment_prod_u)*100/Cint(total_prod_u), 0,-1)
		Set objProd = nothing
		'<!---nsys-publicprofile1-->
		%>		
		
			<script type="text/javascript">
			$(function () {
			    var chart;
			    $(document).ready(function() {
				chart = new Highcharts.Chart({
				    chart: {
					renderTo: 'chartbox',
					type: 'column',
					width: 350,
					height: 250,
					spacingTop:15			
				    },
				    title: {
					text: ''
				    },
				    xAxis: {
					categories: [
					'<%=lang.getTranslated("backend.utenti.detail.table.label.like")%>',
					'<%=lang.getTranslated("frontend.area_user.manage.label.content_comment")&" "&lang.getTranslated("frontend.area_user.manage.label.content_comment2")%>'
					/*<!--nsys-publicprofile2-->*/
					,'<%=lang.getTranslated("frontend.area_user.manage.label.product_comment")&" "&lang.getTranslated("frontend.area_user.manage.label.product_comment2")%>'
					/*<!---nsys-publicprofile2-->*/
					],
					labels: {
						rotation: -20,
						align: 'right',
						style: {
							fontWeight: 'bold'
						}
					}
				    },
				    yAxis: {
					title: {
					    text: ''
					},
					plotLines: [{
					    value: 0,
					    width: 1,
					    color: '#808080'
					}],
					min: 0,
					max:100,
					tickInterval: 20
				    },
				    tooltip: {
					formatter: function() {
					    return ''+
						this.x +': '+ this.y +' %';
					}
				    },
				    legend: {
					enabled: false
				    },
				    series: [{
					data: [
					{
						color: 'blue',
						y: <%=percentual_u%>
					}, {
						color: 'red',
						y: <%=total_news_u%>
					}
					/*<!--nsys-publicprofile2-->*/
					, {
						color: 'green',
						y: <%=total_prod_u%>
					}
					/*<!---nsys-publicprofile2-->*/
					],
					dataLabels: {
						enabled: true,
						color: 'black',
						style: {
						fontWeight: 'bold'
						},
						formatter: function() {
						return this.y +'%';
						}
					}
				    }]
				});
			    });
			    
			});
			</script>
			<div align="left" id="chartbox" style="width:350px;height:237px;border:#c0c0c0 1px solid;overflow: hidden;"></div>			
		
		</div>

		<!--***************************************** TAB INFO *****************************************-->

		<div id="info" style="visibility:hidden;display:none;">
		<div id="profilo-utente">
        	 <h2><%=lang.getTranslated("frontend.header.label.utente_profile_group")%></h2>
             <ul>
                <li><span><%=lang.getTranslated("frontend.area_user.manage.label.username")%> (*)</span></li>
                <li><em><%=strUserName%></em></li>
				<li>
          <script>
            $(function() {
              $(".imgAvatarUser").aeImageResize({height: 50, width: 50});
            });
          </script>        
          <%if(usrHasImg)then%>
					<img class="imgAvatarUser" src="<%=Application("baseroot") & "/common/include/userImage.asp?userID="&id_utente%>" />
					<!--<script>resizeimagesByID('imgAvatarUserProfile', 50);</script>-->
					<%else%>
					<img class="imgAvatarUser" src="<%=Application("baseroot") & "/common/img/unkow-user.jpg"%>"/>
					<%end if%></li>
            <!--<br/>
                <li><span><%'=lang.getTranslated("frontend.area_user.manage.label.email")%> (*)</span></li>
                <li><%'=strEmail%></li>-->
            <br/>
			<li><span><%=lang.getTranslated("frontend.area_user.manage.label.public_profile")%></span></li>
			<li>
				<%if (strComp("1", bolPublic, 1) = 0) then response.Write(lang.getTranslated("portal.commons.yes"))%>
                <%if (strComp("0", bolPublic, 1) = 0) then response.Write(lang.getTranslated("portal.commons.no"))%>
			</li>
			</ul>
            
       <!--******** GESTIONE FIELDS UTENTE PERSONALIZZATI ********-->

        <%
        '********** RECUPERO LA LISTA DI FIELD UTENTE DISPONIBILI
        Dim objUserField, objListUserField, objUserFieldGroup, strPrecFieldgroup, strFieldgroup, fieldMatchValue, hasUserFields
        hasUserFields=false
        On Error Resume Next
        Set objUserField = new UserFieldClass
        Set objListUserField = objUserField.getListUserField(1,"1,3")
        if(objListUserField.count > 0)then
          hasUserFields=true
        end if
        if(Err.number <> 0) then
          hasUserFields=false
        end if                  
                
        strPrecFieldgroup = ""
        fieldMatchValue = ""
                
        Dim userFieldcount
        userFieldcount =1
        if(hasUserFields) then
        for each k in objListUserField
            On Error Resume next
            Set objField = objListUserField(k)

              fieldMatchValue = objUserField.findFieldMatchValue(objField.getID(),id_utente)
          
              select Case objField.getTypeField()
              Case 6,7
                if not(fieldMatchValue = "") AND not(isNull(fieldMatchValue)) then
                  fieldMatchValueArr = split(fieldMatchValue,",")
                  fieldMatchValue = ""
                  fieldMatchValueTmp =""
                  for j=0 to Ubound(fieldMatchValueArr)
                      if not(lang.getTranslated("frontend.area_user.manage.label."&fieldMatchValueArr(j))="") then fieldMatchValueTmp = lang.getTranslated("frontend.area_user.manage.label."&fieldMatchValueArr(j)) else fieldMatchValueTmp=fieldMatchValueArr(j) end if
                      fieldMatchValue = fieldMatchValue & fieldMatchValueTmp&"<br/>"
                  next
                end if              
              Case Else
              End Select 
          
            if(userFieldcount=1) then
              strFieldgroup = objField.getObjGroup().getDescription()
              strPrecFieldgroup = strFieldgroup%>	      
              <h2><%if not(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)="") then response.write(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)) else response.write(strFieldgroup) end if%></h2>
	      <ul>
              <%if(Cint(objField.getTypeField())<>8)then%>
		<li><span><%if not(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())="") then response.write(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())) else response.write(objField.getDescription()) end if%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></li>
              <li><%=fieldMatchValue%></li><br>
          <%end if
          else
                strFieldgroup = objField.getObjGroup().getDescription()
                if(strFieldgroup = strPrecFieldgroup) then
		if(Cint(objField.getTypeField())<>8)then%>
                  <li><span><%if not(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())="") then response.write(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())) else response.write(objField.getDescription()) end if%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></li>
                  <li><%=fieldMatchValue%></li><br>                
                <%end if
          else%>
                  </ul>
                  <h2><%if not(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)="") then response.write(lang.getTranslated("frontend.area_user.manage.label.group."&strFieldgroup)) else response.write(strFieldgroup) end if%></h2>
                  <ul>
		  <%if(Cint(objField.getTypeField())<>8)then%>
                  <li><span><%if not(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())="") then response.write(lang.getTranslated("frontend.area_user.manage.label."&objField.getDescription())) else response.write(objField.getDescription()) end if%><%if(CInt(objField.getRequired())=1) then response.write("&nbsp;(*)")%></span></li>
                  <li><%=fieldMatchValue%></li><br>                
                <%end if
          strPrecFieldgroup = strFieldgroup
                  end if              
              end if
                  
            if(userFieldcount = objListUserField.Count) then response.write("</ul>") end if
              
            userFieldcount=userFieldcount+1

            if(Err.number<>0) then
            response.write(Err.description)
            end if 
        next 
        end if

      Set objListUserField = nothing
      Set objUserField = nothing
      %>
      <!--******** FINE GESTIONE FIELDS UTENTE PERSONALIZZATI ********-->
        
        
			<!--
			<h2><%'=lang.getTranslated("frontend.header.label.iscriz_newsletter")%></h2>
			<ul>
			<li><span><%'=lang.getTranslated("frontend.area_user.manage.label.iscriz_newsletter")%></span></li>
			<li id="profilo-utente-newsletter">
		    <%
			Dim hasNewsletter, objNewsletterTmp
			hasNewsletter = false
			on error Resume Next
			
				Set objListaNewsletter = objNewsletter.getListaNewsletter(1)
				if isObject(objListaNewsletter) AND not(isNull(objListaNewsletter)) AND not (isEmpty(objListaNewsletter)) then
					if(objListaNewsletter.Count > 0) then
						hasNewsletter = true
					end if
				end if
				
			if Err.number <> 0 then
				'response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
			end if	
			
			if(hasNewsletter) then
					dim chechedVal
					for each x in objListaNewsletter.Keys			
						Set objNewsletterTmp = objListaNewsletter(x)
						if not(isNull(objNewsletterUsr)) then
							chechedVal = ""
							if objNewsletterUsr.Exists(x)= true then
								'response.write(objNewsletterTmp.getDescrizione()&"<br/>")
							end if
						end if
						%>		  
						<%Set objNewsletterTmp = nothing
					next%>
			<%end if%>
			</li>		
			</ul>	-->
       </div>
	   </div>

		<!--***************************************** TAB FOTO *****************************************-->

  		<div id="foto" style="visibility:hidden;display:none;">
        <div id="flashcontent">TiltViewer requires JavaScript and the latest Flash player. <a href="http://www.macromedia.com/go/getflashplayer/">Get Flash here.</a></div>
        <script type="text/javascript">
        
          var fo = new SWFObject("<%=Application("baseroot")&"/common/swf/TiltViewer.swf"%>", "viewer", "100%", "400px", "9.0.28", "#000000");			
          
          // TILTVIEWER CONFIGURATION OPTIONS
          // To use an option, uncomment it by removing the "//" at the start of the line
          // For a description of config options, go to: 
          // http://www.airtightinteractive.com/projects/tiltviewer/config_options.html
                                    
          //FLICKR GALLERY OPTIONS
          // To use images from Flickr, uncomment this block
          //fo.addVariable("useFlickr", "true");
          //fo.addVariable("user_id", "48508968@N00");
          //fo.addVariable("tags", "jump,smile");
          //fo.addVariable("tag_mode", "all");
          //fo.addVariable("showTakenByText", "true");			
          
          // XML GALLERY OPTIONS
          // To use local images defined in an XML document, use this block		
          fo.addVariable("useFlickr", "false");
          fo.addVariable("xmlURL", "<%=Application("baseroot")&"/area_user/userphotoxml.asp?userID="&id_utente%>");
          fo.addVariable("maxJPGSize","640");
          
          //GENERAL OPTIONS		
          fo.addVariable("useReloadButton", "false");
          fo.addVariable("columns", "3");
          fo.addVariable("rows", "3");
          //fo.addVariable("showFlipButton", "true");
          //fo.addVariable("showLinkButton", "true");		
          fo.addVariable("linkLabel", "Delete image");
          //fo.addVariable("frameColor", "0xFF0000");
          //fo.addVariable("backColor", "0xDDDDDD");
          //fo.addVariable("bkgndInnerColor", "0xFF00FF");
          //fo.addVariable("bkgndOuterColor", "0x0000FF");				
          //fo.addVariable("langGoFull", "Go Fullscreen");
          //fo.addVariable("langExitFull", "Exit Fullscreen");
          //fo.addVariable("langAbout", "About");				
          
          // END TILTVIEWER CONFIGURATION OPTIONS
          
          fo.addParam("allowFullScreen","true");
          fo.write("flashcontent");			
        </script>      
      </div>
			<%
			Set objUserClass = nothing
			Set objComment = nothing
			Set objUserPreference = nothing
			%>	
		   
<!-- #include file="grid_bottom.asp" -->
</body>
</html>