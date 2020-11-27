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
