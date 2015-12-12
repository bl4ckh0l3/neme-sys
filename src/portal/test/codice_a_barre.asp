<%
Public Function Code39(chaine)
  'V 1.0.0
  'Paramètres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affichée avec la police CODE39.TTF, donne le code barre
  '         * une chaine vide si paramètre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE39.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i
  Code39 = ""
  If Len(chaine) > 0 Then
  'Vérifier si caractères valides
  'Check for valid characters
    For i = 1 To Len(chaine)
      Select Case Asc(Mid(chaine, i, 1))
      Case 32, 36, 37, 43, 45 To 57, 65 To 90
      Case Else
        i = 0
        Exit For
      End Select
    Next
    If i > 0 Then
      Code39 = "*" & chaine & "*"
    End If
  End If
End Function

response.write(Code39("1234567grtyh89"))
%>