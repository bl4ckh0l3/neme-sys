		<!-- #include virtual="/fckeditor/fckeditor.asp" -->			
		<%
		'*************** INIZIALIZZO IL CODICE PER GENERARE GLI EDITOR HTML
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.Width = 600
		oFCKeditor.Height = 300
		oFCKeditor.BasePath = "/fckeditor/"
		'oFCKeditor.Value = 
		oFCKeditor.Create "mail_body"
		%>