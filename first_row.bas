Sub tlacitko_prvy_zaznam()
'
' go to the first row directly
'

If ActiveSheet.Name = "Knihy_L'uboš" Then
    ActiveSheet.Range("$N$4").Select
End If
If ActiveSheet.Name = "Knihy_Žanetka" Then
    ActiveSheet.Range("$N$4").Select
End If
If ActiveSheet.Name = "LP" Or ActiveSheet.Name = "Èasopisy" Then
    ActiveSheet.Range("$B$4").Select
End If
End Sub
