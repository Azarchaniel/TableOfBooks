Sub tlacitko_posl_zaznam()
'
' Makro4 Makro
'function for going to the last row / last book
'it will go to the last row of table (position 999) and then up to the nearest row with value
'Before I was doing it with just Selection.End(xlDown).Select but there was a problem, if there
'was some empty cell
'

If ActiveSheet.Name = "Knihy_L'uboš" Then
    ActiveSheet.Range("K999").Select
    Selection.End(xlUp).Select
End If
If ActiveSheet.Name = "Knihy_Žanetka" Then
    ActiveSheet.Range("K999").Select
    Selection.End(xlUp).Select
End If
If ActiveSheet.Name = "LP" Then
    ActiveSheet.Range("B500").Select
    Selection.End(xlUp).Select
End If
End Sub
