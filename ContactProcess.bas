Attribute VB_Name = "ContactProcess"
Function processContact(numero As String)
' enlever le caractere - (substring) WorksheetFunction.Substitute(cell, " ", ",")
newChaine = WorksheetFunction.Substitute(numero, "-", "")
' compter le nombre de caractere nbcar Len()
nbcar = Len(newChaine)
' si le nb de caracter est 9 cest bon on returne
If nbcar = 9 Then
    processContact = newChaine
End If
' sinon si le nb de caractere est 8 on ajout 6
If nbcar = 8 Then
   processContact = "6" & newChaine
End If


' si non on comment un autre traitment ici
'***********************************
    ' on enleve le +
newChaine2 = WorksheetFunction.Substitute(newChaine, "+", "")
'MsgBox "CHAINE 2 " & newChaine2
' on prend les 3 premier de gauche gauche Left(texte, 7)
troisPremierCar = Left(newChaine2, 3)
' on regarde si cest 237 et que nombre de car > a 8 on retourne nbcar- 3 sinon rien droite Right(texte, 5)
newChaine3 = Right(newChaine2, Len(newChaine2) - 3)
'MsgBox "CHAINE 3 " & newChaine3
If (troisPremierCar = "237" And Len(newChaine3) = 9) Or (troisPremierCar <> "237" And Len(newChaine3) = 9) Then
    processContact = newChaine3
ElseIf troisPremierCar = "237" And Len(newChaine3) = 8 Or (troisPremierCar <> "237" And Len(newChaine3) = 8) Then
    processContact = "6" & newChaine3
End If



        
End Function


Sub ex()
    Debug.Print processContact("99044427")
End Sub
