<div id="backend-header">
  <div id="backend-header-container">
    <div id="backend-header-logo"><a class="logo" href="<%=Application("baseroot")&"/editor/index.asp"%>"><img src="<%=Application("baseroot")&"/editor/img/logo.png"%>" hspace="0" vspace="0" border="0" align="left"></a></div>

    <div id="backend-header-language">
      <%
      Dim objSelLanguage, objTmpSelLanguage, objSelListaLanguage, strActiveLang
      Set objSelLanguage = New LanguageClass
      if isObject(objSelLanguage.getListaLanguageByDesc()) AND not(isEmpty(objSelLanguage.getListaLanguageByDesc())) then
        Set objSelListaLanguage = objSelLanguage.getListaLanguageByDesc()
        objTmpSelLanguage = objSelListaLanguage.Keys		
        for each z in objTmpSelLanguage
        strActiveLang = ""
        if(strComp(z, langEditor.getLAngCode(), 1) = 0) then
          strActiveLang = "-active"
        end if%>
        
        <a title="<%'=langEditor.getTranslated("backend.header.tips_nav_lang")%><%=langEditor.getTranslated("portal.header.label.desc_lang."&z)%>" class="link-lang-fruizione<%=strActiveLang%>" href="<%=Application("baseroot") & "/editor/index.asp?langEditorCode="&z%>"><img src="<%=Application("baseroot")&"/editor/img/flag/flag-"&z&".png"%>" alt="<%=z%>" width="16" height="11" border="0" /><%'=z%></a>
        <%next
        Set objSelListaLanguage = nothing	
      end if
      Set objSelLanguage = nothing
      %>
    </div> 
      
    <div id="backend-header-user">
    <%if not(isEmpty(Session("objCMSUtenteLogged"))) then
      Dim objUsrLoggedTmp, strUsrName
      Set objUsrLoggedTmp = new UserClass
      strUsrName = objUsrLoggedTmp.findUserByID(Session("objCMSUtenteLogged")).getUserName()
      Set objUsrLoggedTmp = nothing%> 
    <strong><%=langEditor.getTranslated("backend.header.utente")%>:</strong>&nbsp;<%=strUsrName%>&nbsp;&nbsp;<input type="button" class="buttonLogOff" hspace="2" vspace="0" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.header.logoff")%>" onclick="javascript: document.location.href='<%=Application("baseroot")&"/common/include/LogOFF.asp"%>';" />
    <%end if%>
    </div>
  </div>
</div>