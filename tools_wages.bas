Attribute VB_Name = "Module1"
Function tdl(sbm)
    
    If sbm < 62000 Then
        tdl = 0
    End If
    
    If (sbm > 62000) And (sbm < 75001) Then
        tdl = 250
    End If
    
    If (sbm > 75001) And (sbm < 100001) Then
        tdl = 500
    End If
    
    If (sbm > 100001) And (sbm < 125001) Then
        tdl = 750
    End If
    
    If (sbm > 125001) And (sbm < 150001) Then
        tdl = 1000
    End If
    
    If (sbm > 150001) And (sbm < 200001) Then
        tdl = 1250
    End If
    
    If (sbm > 200001) And (sbm < 250001) Then
        tdl = 1500
    End If
    
    If (sbm > 250001) And (sbm < 300001) Then
        tdl = 2000
    End If
    
    If (sbm > 300001) And (sbm < 500001) Then
        tdl = 2250
    End If
    
    If sbm > 500001 Then
        tdl = 2500
    End If
    
End Function



Function rdv(salbase)
    
    If salbase < 50000 Then
        rdv = 0
    End If
    
    If (salbase > 50000) And (salbase <= 100000) Then
        rdv = 750
    End If
    
    If (salbase > 100000) And (salbase <= 200000) Then
        rdv = 1950
    End If
    
    If (salbase > 200000) And (salbase <= 300000) Then
        rdv = 3250
    End If
    
    If (salbase > 300000) And (salbase <= 400000) Then
        rdv = 4550
    End If
    
    If (salbase > 400000) And (salbase <= 500000) Then
        rdv = 5850
    End If
    
    If (salbase > 500000) And (salbase <= 600000) Then
        rdv = 7150
    End If
    
    If (salbase > 600000) And (salbase <= 700000) Then
        rdv = 8450
    End If
    
     If (salbase > 700000) And (salbase <= 800000) Then
        rdv = 9750
    End If
    
    If (salbase > 800000) And (salbase <= 900000) Then
        rdv = 11050
    End If
    
    
    If (salbase > 900000) And (salbase <= 1000000) Then
        rdv = 12350
    End If
    
    If (salbase > 1000000) Then
        rdv = 13000
    End If
    
   
End Function



Function irpp(ByVal baseIrpp, Optional ByVal salbase)
   
    
   ' si on veut que le test sois fait sur le salbase
   
        If salbase <= 62000 Or baseIrpp <= 62000 Then
            irpp = 0
            
        ElseIf baseIrpp <= 166667 Then
            irpp = baseIrpp * 0.1
            
        ElseIf baseIrpp <= 250000 Then
            irpp = (baseIrpp - 166667) * 0.15 + 16667
            
        ElseIf baseIrpp < 416667 Then
            irpp = (baseIrpp - 250000) * 0.25 + 29167
            
        ElseIf baseIrpp > 416667 Then
            irpp = (baseIrpp - 416667) * 0.35 + 70833
        End If

    
    
End Function
