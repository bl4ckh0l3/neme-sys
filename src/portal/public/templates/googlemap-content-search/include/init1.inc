<%
Dim News, objListaNews, objListaTargetCat, objListaTargetLang
Dim strGerarchia, totPages, numPage, stato

Set News = New NewsClass
strGerarchia = request("gerarchia")
numPage = 1
stato = 1

if not(isNull(request("page"))) AND not(request("page") = "") then
	numPage = request("page")
end if

if(request("is_preview_content")) then
	stato = null
end if
%>
<!-- #include virtual="/common/include/setTemplateTargetList.inc" -->
<%
Dim bolHasObj
bolHasObj = false

Set objContentF = new ContentFieldClass
Set objListPairKeyValue = Server.CreateObject("Scripting.Dictionary")	
Set objListPairKeyValueRules = Server.CreateObject("Scripting.Dictionary")	

On Error Resume Next
Set obiListCF = objContentF.getListContentField(1)

for each j in obiListCF
  tmpCfDesc = obiListCF(j).getDescription()

  Select Case tmpCfDesc
    Case "contratto"  
      Set objCON = obiListCF(j)
      idcon = objCON.getID()
    Case "categoria"  
      Set objCAT = obiListCF(j)
      idcat = objCAT.getID()
    Case "tipologia"  
      Set objTYP = obiListCF(j)
      idtyp = objTYP.getID()
    Case "tipo-proprieta"  
      Set objTPP = obiListCF(j)
      idtpp = objTPP.getID()
    Case "stato"  
      Set objSTA = obiListCF(j)
      idsta = objSTA.getID()
    Case "prezzo"  
      Set objPRZ = obiListCF(j)
      idprc = objPRZ.getID()
      objListPairKeyValueRules.add idprc, "prezzo"
    Case "superficie"  
      Set objSUP = obiListCF(j)
      idsup = objSUP.getID()
      objListPairKeyValueRules.add idsup, "superficie"
    Case "locali"  
      Set objLOC = obiListCF(j)
      idloc = objLOC.getID()
      objListPairKeyValueRules.add idloc, "locali"
    Case "bagni"  
      Set objBAT = obiListCF(j)
      idbat = objBAT.getID()
    Case "riscaldamento"  
      Set objRIS = obiListCF(j)
      idris = objRIS.getID()
    Case "accessori"  
      Set objACC = obiListCF(j)
      idacc = objACC.getID()
    Case "regione-provincia"  
      Set objRPV = obiListCF(j)
      idrpv = objRPV.getID()
      objListPairKeyValueRules.add idrpv, "regione-provincia"
    Case else               
  End Select
next
Set obiListCF = nothing

if(Err.number<>0)then
'response.write(Err.description)
end if

bolHasFieldsFilter = false
bolHasSessionActive = false
bolHasFilterSearchActive = false
bolHasGeoSearchActive = false
strParamPagFilter=""

if (Instr(1, typename(Session("geolocalsearchpoly")), "Dictionary", 1) > 0) then
    bolHasSessionActive = true
end if

On Error Resume Next
'****************** verifico se è stata fatta una ricerca dei field per contenuto e in tal caso recupero solo le news filtrate
if(request("fields_filter")="1")then
  if not(bolHasSessionActive)then
    Set objGeolocalSearch = Server.CreateObject("Scripting.Dictionary") 
    Set Session("geolocalsearchpoly") = objGeolocalSearch  
  end if

  bolHasFilterSearchActive = true
  Session("geolocalsearchpoly").item("fields_filter_active") = "1"

  if (Request.Form.Count>0)then      
    For Each y In Request.Form
      if (Instr(1, y, "field_", 1) > 0) then
        if(Trim(Request.Form(y))<>"") then
          tmpKey = Mid(y,Instr(1, y, "__", 1)+2)
          tmpValue = Request.Form(y)
          'response.write("tmpKey:"&tmpKey&" - tmpValue:"&tmpValue&"<br>")
          objListPairKeyValue.add tmpKey,tmpValue
          strParamPagFilter=strParamPagFilter&y&"="&Request.Form(y)&"&"
        end if
      end if
    Next 

    strParamPagFilter=strParamPagFilter&"fields_filter=1"
 
    if(objListPairKeyValue.Count > 0)then
      Set Session("geolocalsearchpoly").item("objListPairKeyValue") =  objListPairKeyValue      
      Session("geolocalsearchpoly").item("strParamPagFilter") =  strParamPagFilter
      'response.write("Session objListPairKeyValue.count:"& Session("geolocalsearchpoly").item("objListPairKeyValue").Count &"<br>")
    end if
    'response.write("strParamPagFilter:"&strParamPagFilter&" - objListPairKeyValue.Count > 0: "& (objListPairKeyValue.Count > 0) &"<br>")
    if(objListPairKeyValue.Count > 0)then
      bolHasFieldsFilter = true
    end if
  end if  
else
  '************ verifico se c'è una ricerca in sessione e riattivo i filtri
  if(bolHasSessionActive)then
    if (Instr(1, typename(Session("geolocalsearchpoly").item("objListPairKeyValue")), "Dictionary", 1) > 0) then
      Set objListPairKeyValue = Session("geolocalsearchpoly").item("objListPairKeyValue")
      strParamPagFilter = Session("geolocalsearchpoly").item("strParamPagFilter") 
      'response.write("strParamPagFilter:"&strParamPagFilter&"<br>")
      for each u in objListPairKeyValue
          'response.write("tmpKey Session:"&u&" - tmpValue Session:"&objListPairKeyValue(u)&"<br>")
        if(Trim(objListPairKeyValue(u))="")then
          objListPairKeyValue.remove(u)  
        end if
      next
      bolHasFieldsFilter = true
    end if
    
    tmp_fields_filter_active = Session("geolocalsearchpoly").item("fields_filter_active")
    
    'response.write("tmp_fields_filter_active:"&tmp_fields_filter_active&"<br>")
    
    if(tmp_fields_filter_active = "1")then
      bolHasFilterSearchActive = true
    end if
    'response.write("bolHasFilterSearchActive:"& bolHasFilterSearchActive &"<br>")
  end if
end if

if(Err.number <> 0) then
  'response.write(Err.description)
  bolHasFieldsFilter = false
end if 


Dim objLocaliz, objPoint
Set objLocaliz = New LocalizationClass
Set objListPoint = Server.CreateObject("Scripting.Dictionary")

'**** verifico se è stata lanciata una ricerca e imposto nell'oggetto in sessione il parametro della ricerca attiva
if(request("search_active")="1")then
  Session("geolocalsearchpoly").Item("search_active") = "1"
end if

if (bolHasSessionActive) then
  if(Session("geolocalsearchpoly")("search_active")="1")then
    bolHasGeoSearchActive = true
  else
    Session("geolocalsearchpoly").remove("type")
    Session("geolocalsearchpoly").remove("current_overlay")
    Session("geolocalsearchpoly").remove("last_selection")
    Session("geolocalsearchpoly").remove("vertices")
    Session("geolocalsearchpoly").remove("center")
    Session("geolocalsearchpoly").remove("radius")   
  end if
end if

'response.write("bolHasGeoSearchActive:"& bolHasGeoSearchActive &"<br>")

if(bolCatContainNews) AND not(isNull(objListaTargetCat)) then
  if(bolHasFilterSearchActive)then
    on error Resume Next
    Set objListaNews = News.findNewsSlimCached(null, null, null, null, objListaTargetCat, objListaTargetLang, null, null, stato, order_by,false,true)	
    
    if(objListaNews.Count > 0) then		
      bolHasObj = true
    end if
    
    if Err.number <> 0 then
      bolHasObj = false
    end if
  end if


  'response.write("bolHasObj:"& bolHasObj &"<br>")
  
  if(bolHasObj)then
    for each t in objListaNews
      keepContent=true
      on error Resume Next    
      '*************** verifico se esistono fields per questo elemento e li filtro in base alla ricerca      
      Set objListContentFieldTmp = objContentF.getListContentField4ContentActiveCached(t)      
      if(objListContentFieldTmp.Count > 0)then        
        if(bolHasFieldsFilter)then 
          for each j in objListPairKeyValue
            if(objListContentFieldTmp.Exists(Clng(j))=false)then
              objListaNews.remove(t)
              keepContent=false
              exit for
            end if               
              
            Select Case objListPairKeyValueRules(Clng(j))
              Case "prezzo"
                hasLeft = false
                hasRight=false
                tmpcfVals = CDbl(objListContentFieldTmp(Clng(j)).getSelValue())
                tmpprcVals = Split(objListPairKeyValue(j), "x", -1, 1)              
            
                if(Trim(tmpprcVals(0))="" OR Trim(tmpprcVals(1))="")then
                  prcCheck = Trim(objListPairKeyValue(j))
                
                  if(Right(prcCheck,1)="x")then
                      prcsxVal = CDbl(Trim(tmpprcVals(0)))
                      hasLeft = true
                  elseif(Left(prcCheck,1)="x")then
                    prcCheck = Trim(Replace(prcCheck, "x", "", 1, -1, 1))
                    tmpprcVals(1) = prcCheck
                    prcdxVal = CDbl(Trim(tmpprcVals(1)))
                    hasRight=true                      
                  end if
                else
                  prcsxVal = CDbl(Trim(tmpprcVals(0)))
                  prcdxVal = CDbl(Trim(tmpprcVals(1)))
                  hasLeft = true
                  hasRight=true                   
                end if
              
                'response.write("<br>hasLeft:"& hasLeft &" -hasRight:"& hasRight & " -prcsxVal:"&prcsxVal&" -prcdxVal:"&prcdxVal&"<br>")
              
                if(hasLeft)then
                  'response.write("<br>tmpcfVals:"& tmpcfVals &" -typename:"& typename(tmpcfVals)&"<br>")                  
                
                  if(tmpcfVals < prcsxVal)then
                    objListaNews.remove(t)
                    keepContent=false
                    exit for
                  end if 
                end if 
              
                if(hasRight)then
                  'response.write("<br>tmpcfVals:"& tmpcfVals &" -typename:"& typename(tmpcfVals)&"<br>")  
                
                  if(tmpcfVals > prcdxVal)then
                    objListaNews.remove(t)
                    keepContent=false
                    exit for
                  end if     
                end if                 
              Case "superficie"
                hasLeft = false
                hasRight=false
                tmpcfVals = Clng(objListContentFieldTmp(Clng(j)).getSelValue())
                tmpsupVals = Split(objListPairKeyValue(j), "x", -1, 1)              
            
                if(Trim(tmpsupVals(0))="" OR Trim(tmpsupVals(1))="")then
                  supCheck = Trim(objListPairKeyValue(j))
                
                  if(Right(supCheck,1)="x")then
                      supsxVal = Clng(Trim(tmpsupVals(0)))
                      hasLeft = true
                  elseif(Left(supCheck,1)="x")then
                    supCheck = Trim(Replace(supCheck, "x", "", 1, -1, 1))
                    tmpsupVals(1) = supCheck
                    supdxVal = Clng(Trim(tmpsupVals(1)))
                    hasRight=true                      
                  end if
                else
                  supsxVal = Clng(Trim(tmpsupVals(0)))
                  supdxVal = Clng(Trim(tmpsupVals(1)))
                  hasLeft = true
                  hasRight=true                   
                end if
              
                if(hasLeft)then
                  if(tmpcfVals < supsxVal)then
                    objListaNews.remove(t)
                    keepContent=false
                    exit for
                  end if 
                end if 
              
                if(hasRight)then
                  if(tmpcfVals > supdxVal)then
                    objListaNews.remove(t)
                    keepContent=false
                    exit for
                  end if     
                end if                   
              Case "locali"
                hasLeft = false
                hasRight=false
                tmpcfVals = Cint(objListContentFieldTmp(Clng(j)).getSelValue())
                tmplocVals = Split(objListPairKeyValue(j), "x", -1, 1)              
            
                if(Trim(tmplocVals(0))="" OR Trim(tmplocVals(1))="")then
                  locCheck = Trim(objListPairKeyValue(j))
                
                  if(Right(locCheck,1)="x")then
                      locsxVal = Cint(Trim(tmplocVals(0)))
                      hasLeft = true
                  elseif(Left(locCheck,1)="x")then
                    locCheck = Trim(Replace(locCheck, "x", "", 1, -1, 1))
                    tmplocVals(1) = locCheck
                    locdxVal = Cint(Trim(tmplocVals(1)))
                    hasRight=true                      
                  end if
                else
                  locsxVal = Cint(Trim(tmplocVals(0)))
                  locdxVal = Cint(Trim(tmplocVals(1)))
                  hasLeft = true
                  hasRight=true                   
                end if
              
                if(hasLeft)then
                  if(tmpcfVals < locsxVal)then
                    objListaNews.remove(t)
                    keepContent=false
                    exit for
                  end if 
                end if 
              
                if(hasRight)then
                  if(tmpcfVals > locdxVal)then
                    objListaNews.remove(t)
                    keepContent=false
                    exit for
                  end if     
                end if           
              Case "regione-provincia"
                regprovval = objListContentFieldTmp(Clng(j)).getSelValue()
                regprovsel = objListPairKeyValue(j)
                'response.write("regprovsel:"&regprovsel)
                regprovsel = Left(regprovsel,Instr(1, regprovsel, "_", 1)-1)
                'response.write(" - regprovsel intr:"&regprovsel&" - regprovval:"&regprovval&" - instring:"& Instr(1, regprovval, regprovsel, 1) &"<br>")
                if (Instr(1, regprovval, regprovsel, 1) <= 0) then
                    objListaNews.remove(t)
                    keepContent=false
                end if
              Case else
                'response.write("<br>entrato nel caso case else")
                'response.write("j:"&j &" -objListPairKeyValue(j):"&objListPairKeyValue(j)&"<br>")
              
                select Case objListContentFieldTmp(Clng(j)).getTypeField()
                Case 4,5
                  selected = false
                  spitValues = objListContentFieldTmp(Clng(j)).getSelValue()
                  'response.write("spitValues(4,5,6):"&spitValues&"<br>")
                  if not(spitValues = "") then
                    spitValues = Split(spitValues,",")
                    for each x in spitValues                    
                      if not(objListPairKeyValue(j) = "") then
                        spitMatchValues = Split(objListPairKeyValue(j),",")
                        for w=0 to Ubound(spitMatchValues)
                          if(strComp(Trim(spitMatchValues(w)), Trim(x), 1) = 0) then
                            selected=true
                            exit for
                          end if
                        next
                      end if		
                    next 
                  end if
                    
                  if not(selected)then
                    objListaNews.remove(t)
                    keepContent=false
                    exit for
                  end if            
                Case Else    
                  'response.write("objListContentFieldTmp(Clng(j)).getSelValue():"&objListContentFieldTmp(Clng(j)).getSelValue()&"<br>")          
                  if(strComp(objListContentFieldTmp(Clng(j)).getSelValue(), objListPairKeyValue(j), 1) <> 0)then
                    objListaNews.remove(t)
                    keepContent=false
                    exit for
                  end if  
                End Select                  
            End Select
          next  
        end if
      end if
                        
      if(Err.number <> 0) then
        keepContent = false
        'response.write(Err.description)
      end if
      Set objListContentFieldTmp = nothing    
    
 
      if(keepContent)then             
        on error Resume Next   
        '*************** verifico se esiste la geolocalizzazione per questo elemento   
        Set points = objLocaliz.findPointByElement(objListaNews(t).getNewsID(), 1)
        if (Instr(1, typename(points), "Dictionary", 1) > 0) then
          'se esiste un poligono/cerchio impostato dall utente come base di ricerca verifico se il punto è incluso nel poligono/cerchio
          'l oggetto in sessione è un Dictionary con n chiavi, a seconda del tipo di poligono alcune chiavi saranno null e altre valorizzate:
          'type: 1=poligono, 2=cerchio
          'vertices: dictionary contenente i singoli vertici del poligono
          'center: il punto di centro del cerchio
          'radius: il raggio in metri del cerchio
          bolAddAllPoints = true
          if (bolHasGeoSearchActive) then
            bolAddAllPoints = false              
            if(Session("geolocalsearchpoly")("type")=1)then
              Set objListVertices = objLocaliz.convertVertices(Session("geolocalsearchpoly")("vertices"))
              for each xLocal in points.Items
                'response.write("xLocal:"&xLocal.getLatitude()&","&xLocal.getLongitude()&"<br>")            
                if(objLocaliz.isPointInPolygon(xLocal, objListVertices))then
                'if(objLocaliz.coordinate_is_inside_polygon(xLocal.getLatitude(), xLocal.getLongitude(), objListVertices.keys, objListVertices.items))then
                    'response.write("add xLocal<br>")
                    objListPoint.add xLocal, xLocal.getInfo()
                else
                  'response.write("removing t:"&t&"<br>")
                  objListaNews.remove(t)
                  keepContent=false
                  exit for                
                end if
              next
              Set objListVertices = nothing
            elseif(Session("geolocalsearchpoly")("type")=2)then
              Set objCenter = objLocaliz.convertCenter(Session("geolocalsearchpoly")("center"))        
              for each xLocal in points.Items
                'response.write("xLocal:"&xLocal.getLatitude()&","&xLocal.getLongitude()&" - radius:"&Session("geolocalsearchpoly")("radius")&"<br>")           
                if(objLocaliz.IsPointInCircleOnEarthSurface(xLocal, objCenter, Session("geolocalsearchpoly")("radius")))then
                    'response.write("add xLocal<br>")
                    objListPoint.add xLocal, xLocal.getInfo()
                else
                  'response.write("removing t:"&t&"<br>")
                  objListaNews.remove(t)
                  keepContent=false
                  exit for                
                end if
              next 
              Set objCenter = nothing
            end if 
          end if
          
          if(bolAddAllPoints)then
            for each xLocal in points.Items
              objListPoint.add xLocal, xLocal.getInfo()   
            next
          end if
        else      
          if (bolHasGeoSearchActive)then
            objListaNews.remove(t) 
            keepContent=false         
          end if
        end if
        Set points = nothing
        
        if Err.number <> 0 then
          keepContent=false
          'response.write(Err.description&"<br>")
        end if 
      end if
      
      if(keepContent)then
        if not(lang.getTranslated(objListaNews(t).getMetaDescription()) = "") then
          metaDescription = metaDescription & " " & lang.getTranslated(objListaNews(t).getMetaDescription())
        else
          if not(objListaNews(t).getMetaDescription()="")then
            metaDescription = metaDescription & " " & objListaNews(t).getMetaDescription()
          end if
        end if
        
        if not(lang.getTranslated(objListaNews(t).getMetaKeyword()) = "") then
          metaKeyword = metaKeyword & " " & lang.getTranslated(objListaNews(t).getMetaKeyword())
        else
          if not(objListaNews(t).getMetaKeyword()="")then
            metaKeyword = metaKeyword & " " & objListaNews(t).getMetaKeyword()
          end if
        end if
        
        'if not(lang.getTranslated(objListaNews(t).getPageTitle()) = "") then
        ' pageTemplateTitle = pageTemplateTitle & " " & lang.getTranslated(objListaNews(t).getPageTitle())
        'else
        '  if not(objListaNews(t).getPageTitle()="")then
        '    pageTemplateTitle = pageTemplateTitle & " " & objListaNews(t).getPageTitle()
        '  end if
        'end if 
      end if
    next

    metaDescription = Trim(metaDescription)
    metaKeyword = Trim(metaKeyword)
    'pageTemplateTitle = Trim(pageTemplateTitle)  
      
    Dim objSelNews, newsCounter, iIndex, objTmpNews, FromNews, ToNews, Diff
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
  end if
end if

if (pageTemplateTitle = "") then
  pageTemplateTitle = lang.getTranslated("frontend.page.title")
end if  

Response.Charset="UTF-8"
Session.CodePage  = 65001

Set objLocaliz = nothing 


Set objCategory = new CategoryClass
Set objPageTempl = new Page4TemplateClass
Set objTemplate = new TemplateClass
Set objMenuFruizione = new MenuClass
bolHasDetailLink = false
On Error Resume Next							         
Set objCategoriaTmp = objCategory.findExsitingCategoriaByGerarchia(strGerarchia)
if not(isNull(objCategoriaTmp)) then
  Set objTemplateSelected = objTemplate.findTemplateByID(objCategoriaTmp.findLangTemplateXCategoria(lang.getLangCode(),true))
  bolHasDetailLink = true
else
  bolHasDetailLink = false                  
end if
if(Err.number <>0) then
  bolHasDetailLink = false
end if 

' recupero url corrente per definire path a confirm.asp
tmpurl = request.ServerVariables("URL")
tmpurl = Mid(tmpurl,1,InStrRev(tmpurl,"/",-1,1))
%>