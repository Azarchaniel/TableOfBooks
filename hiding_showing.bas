Sub Makro6()
'
' Hiding / showing the columns of authors
'

'
    Dim stlpce As Range
    Dim x As String
 
    
    If ActiveSheet.Name = "Knihy_L'uboš" Then
        Worksheets("Knihy_L'uboš").Activate
        x = "Knihy_L'uboš"
    End If
    If ActiveSheet.Name = "Knihy_Žanetka" Then
        Worksheets("Knihy_Žanetka").Activate
        x = "Knihy_Žanetka"
    End If
    
    Set stlpce = ThisWorkbook.Worksheets(x).Range("C3:J1000")   'select columns of authors
    
    If stlpce.EntireColumn.Hidden = False Then 'if they are not hidden,
        stlpce.EntireColumn.Hidden = True       'hide them and
        Worksheets(x).Buttons("Button 5").Caption = "Show"  'change label of button
    ElseIf stlpce.EntireColumn.Hidden = True Then   'if they are hidden
        stlpce.EntireColumn.Hidden = False          'show them
        Worksheets(x).Buttons("Button 5").Caption = "Hide"
    End If
    
End Sub
