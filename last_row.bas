Sub tlacitko_posl_zaznam()
'
' function for going to the last row / last book
'

If ActiveSheet.Name = "Knihy_L'uboš" Then
    ActiveSheet.Range("$N$3499").Select 'select last row
    Selection.End(xlUp).Select          'then go up to the nearest value
    ActiveCell.Offset(1, 0).Select      'and go one row down
End If
If ActiveSheet.Name = "Knihy_Žanetka" Then
    ActiveSheet.Range("$N$3499").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
End If
If ActiveSheet.Name = "LP" Or ActiveSheet.Name = "Èasopisy" Then
    ActiveSheet.Range("$B$500").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
End If
End Sub
