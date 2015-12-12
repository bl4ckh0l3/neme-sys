		
    <div align="left" id="contenuti">
		<%
		'************** codice per la lista news e paginazione	
		if(bolHasObj) then%>
				<!--<div id="torna"><a href="<%'=Application("baseroot") & "/common/include/feedRSS.asp?gerarchia="&strGerarchia%>" target="_blank"><img src="<%'=Application("baseroot")&"/common/img/rss_image.gif"%>" vspace="3" hspace="3" border="0" align="right" alt="RSS"></a></div>
				<br/>-->	
				<%Dim objSelNews, newsCounter, iIndex, objTmpNews, FromNews, ToNews, Diff
				iIndex = objListaNews.Count
				FromNews = ((numPage * elem_x_page) - elem_x_page)
				Diff = (iIndex - ((numPage * elem_x_page)-1))
				if(Diff < 1) then
					Diff = 1
				end if
				
				ToNews = iIndex - Diff
				
				totPages = iIndex\elem_x_page
				if(totPages < 1) then
					totPages = 1
				elseif((iIndex MOD elem_x_page <> 0) AND not ((totPages * elem_x_page) >= iIndex)) then
					totPages = totPages +1	
				end if		
						
				objTmpNews = objListaNews.Items		
				for newsCounter = FromNews to ToNews
					Set objSelNews = objTmpNews(newsCounter)%>
					<div><p class="title_contenuti"><%=objSelNews.getTitolo()%></p>
					<%if (Len(objSelNews.getAbstract1()) > 0) then response.Write(objSelNews.getAbstract1()) end if
               			if (Len(objSelNews.getAbstract2()) > 0) then response.Write(objSelNews.getAbstract2()) end if
               			if (Len(objSelNews.getAbstract3()) > 0) then response.Write(objSelNews.getAbstract3()) end if
               			if (Len(objSelNews.getTesto()) > 0) then response.Write(objSelNews.getTesto()) end if%>
					</div><p class="line"></p>
					<%Set objSelNews = nothing
				next%>
			  	<div>
				<%if(totPages > 1) then%>
					<%call PaginazioneFrontend(totPages, numPage, strGerarchia, Application("controller_page"), "")
				end if%>
				</div>
		<%else
			response.Write("<br/><br/><div align=""center""><strong>"& lang.getTranslated("portal.commons.templates.label.page_in_progress")&"</strong></div>")
		end if%>
		<div id="profilo-utente">
		<form action="<%=Application("baseroot") &Application("dir_upload_templ")&"contactus/confirm.asp"%>" method="post" name="form_send_mail" onSubmit="return sendMail();">
		  <input type="hidden" name="gerarchia" value="<%=strGerarchia%>">
		  <input type="hidden" name="mailTo" value="<%=Application("mail_receiver")%>">
		  <input type="hidden" name="sent_captchacode" value="">
      <input type="hidden" name="sent_recaptcha_challenge_field" value="">
      <input type="hidden" name="sent_recaptcha_response_field" value="">

		 <!--<h2><%'=lang.getTranslated("frontend.template_contatti.label.testo_intro_mail")%></h2>-->
		 <p><%=lang.getTranslated("frontend.template_contatti.label.testo_intro_mail2")%></p>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.nome")%> (*)</span></li>
			<li><input type="text" name="nome" value="" /></li>
		</ul>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.cognome")%> (*)</span></li>
			<li><input type="text" name="cognome" value="" /></li>
		</ul>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.email")%> (*)</span></li>
			<li><input type="text" name="email" value="" /></li>
		</ul>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.telefono")%> (*)</span></li>
			<li><input type="text" name="telefono" value="" /></li>
		</ul>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.indirizzo")%></span></li>
			<li><input type="text" name="indirizzo" value="" /></li>
		</ul>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.cap_city")%></span></li>
			<li><input type="text" name="zipcode" value="" />&nbsp;<input type="text" name="citta" value="" /></li>
		</ul>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.nazione")%> (*)</span></li>
			<li><input type="text" name="nazione" value="" /></li>
		</ul>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.testo_mail")%></span></li>
			<li><textarea name="testo" rows="3"></textarea></li>
		</ul>
		 <ul>
			<li><span><%=lang.getTranslated("frontend.template_contatti.label.info_privacy")%> (*)</span></li>
			<li><textarea name="testo_privacy" rows="3"><%=lang.getTranslated("frontend.template_contatti.label.info_privacy_law")%></textarea></li>
		</ul>
		 <ul>
			<li><input type="checkbox" name="acceptPrivacy" value="1" hspace="0" vspace="0"><%=lang.getTranslated("frontend.template_contatti.label.privacy_accept")%></li>
		</ul>
		 <ul>
			<li>    
      <%
			  if(request("captcha_err") = 1) then
				response.write("<span  class=imgError>"&lang.getTranslated("frontend.template_contatti.label.wrong_captcha_code") & "</span><br/>")
			  end if
        
			  if(Application("use_recaptcha") = 0) then%>
          <br/><img id="imgCaptcha" src="<%=Application("baseroot")&"/common/include/captcha/base_captcha.asp"%>" />&nbsp;&nbsp;<input name="captchacode" type="text" id="captchacode" />
          <br/><a href="javascript:void(0)" onclick="RefreshImage('imgCaptcha')"><%'=lang.getTranslated("frontend.template_contatti.label.change_captcha_img")%></a>
        <%else%>
          <br/><%=recaptcha_challenge_writer(Application("recaptcha_pub_key"))%>
        <%end if%>
        </li>
		</ul>    
		 <ul>
			<li><br/><input type="submit" name="submit" value="<%=lang.getTranslated("frontend.template_contatti.button.send.label")%>" vspace="0" align="absmiddle">&nbsp;<input type="reset" name="reset" value="<%=lang.getTranslated("frontend.template_contatti.button.cancel.label")%>" vspace="0" align="absmiddle"></li>
		</ul>
		</div>
		</form>
		</div>