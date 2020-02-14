Sub tlacitko_posl_zaznam()
'
' function for going to the last row / last book
'

If ActiveSheet.Name = "Knihy_L'uboš" Or ActiveSheet.Name = "Knihy_Žanetka" Then
    ActiveSheet.Range("$N$2499").Select 'select last row
    Selection.End(xlUp).offset(1, 0).Select          'then go up to the nearest value and go one row down
End If
If ActiveSheet.Name = "LP" Or ActiveSheet.Name = "Časopisy" Then
    ActiveSheet.Range("$B$499").Select
    Selection.End(xlUp).offset(1, 0).Select
End If
End Sub
