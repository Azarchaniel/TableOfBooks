Sub tlacitko_prvy_zaznam()
'
' Makro4 Makro
'go to the first book directly
'

If ActiveSheet.Name = "Knihy_L'uboš" Then
    ActiveSheet.Range("K3").Select
End If
If ActiveSheet.Name = "Knihy_Žanetka" Then
    ActiveSheet.Range("K3").Select
End If
If ActiveSheet.Name = "LP" Then
    ActiveSheet.Range("B3").Select
End If
End Sub
