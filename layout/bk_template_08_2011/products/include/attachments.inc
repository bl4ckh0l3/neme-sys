		<%
		Set objListaFile = New File4ProductsClass
		if not(isNull(objListaFile.getFilePerProdotto(id_prodotto))) AND not(isEmpty(objListaFile.getFilePerProdotto(id_prodotto))) then
			Set objListaFilePerProdotto = objListaFile.getFilePerProdotto(id_prodotto)
			
			if not(isEmpty(objListaFilePerProdotto)) then
				' LEGENDA TIPI FILE
				'1 = img small
				'2 = img big
				'3 = pdf
				'4 = audio-video
				'5 = others%>				
				<%
				' Lista label tipi file
				Dim hasBigImg, hasPdf, hasAudioVideo, hasOthers
				hasBigImg = false
				hasPdf = false
				hasAudioVideo = false
				hasOthers = false
				
				for each xObjFile in objListaFilePerProdotto
					Set objFileXProd = objListaFilePerProdotto(xObjFile)					
					
					select case objFileXProd.getFileTypeLabel()
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
					Set objFileXProd = nothing	
				next
				
				if (cbool(hasBigImg)) then response.write("<br/><strong>"&lang.getTranslated("frontend.file_allegati.label.key_img_big")&"</strong><br/><br/>") end if
				' Lista immagini grandi
				for each xObjFile in objListaFilePerProdotto
					Set objFileXProd = objListaFilePerProdotto(xObjFile)					
					if(objFileXProd.getFileTypeLabel() = 2) then%>
						<a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popupProd.asp?id_allegato="&objFileXProd.getFileID()%>','popupallegati',400,400,100,100)"><%=objFileXProd.getFileName()%></a><br><br>					
						<%objListaFilePerProdotto.remove(xObjFile)
					end if
					Set objFileXProd = nothing	
				next
				
				if (cbool(hasAudioVideo)) then response.write("<br/><strong>"&lang.getTranslated("frontend.file_allegati.label.key_audio_video")&"</strong><br/><br/>") end if
				' Lista audiovideo
				for each xObjFile in objListaFilePerProdotto
					Set objFileXProd = objListaFilePerProdotto(xObjFile)					
					if(objFileXProd.getFileTypeLabel() = 4) then%>
						<a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popupProd.asp?id_allegato="&objFileXProd.getFileID()%>','popupallegati',400,400,100,100)"><%=objFileXProd.getFileName()%></a><br><br>					
						<%objListaFilePerProdotto.remove(xObjFile)
					end if
					Set objFileXProd = nothing	
				next
				
				if (cbool(hasPdf)) then response.write("<br/><strong>"&lang.getTranslated("frontend.file_allegati.label.key_pdf")&"</strong><br/><br/>") end if
				' Lista pdf
				for each xObjFile in objListaFilePerProdotto
					Set objFileXProd = objListaFilePerProdotto(xObjFile)					
					if(objFileXProd.getFileTypeLabel() = 3) then%>
						<a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popupProd.asp?id_allegato="&objFileXProd.getFileID()%>','popupallegati',400,400,100,100)"><%=objFileXProd.getFileName()%></a><br><br>					
						<%objListaFilePerProdotto.remove(xObjFile)
					end if
					Set objFileXProd = nothing	
				next
				
				if (cbool(hasOthers)) then response.write("<br/><strong>"&lang.getTranslated("frontend.file_allegati.label.key_others_doc")&"</strong><br/><br/>") end if
				' Lista others documents
				for each xObjFile in objListaFilePerProdotto
					Set objFileXProd = objListaFilePerProdotto(xObjFile)					
					if(objFileXProd.getFileTypeLabel() = 5) then%>
						<a href="javascript:openWin('<%=Application("baseroot")&"/common/include/popupProd.asp?id_allegato="&objFileXProd.getFileID()%>','popupallegati',400,400,100,100)"><%=objFileXProd.getFileName()%></a><br><br>					
						<%objListaFilePerProdotto.remove(xObjFile)
					end if
					Set objFileXProd = nothing	
				next				
			end if
			Set objListaFilePerProdotto = nothing
		end if%>