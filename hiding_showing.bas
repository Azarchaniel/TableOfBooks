Sub Makro6()
'
' Hiding / showing the columns of authors, ilustrator etc.
'

    Dim stlpce As Range
    Dim stlpce2 As Range
    Dim x As String
 
    If ActiveSheet.Name = "Knihy_L'uboš" Then
        Worksheets("Knihy_L'uboš").Activate
        x = "Knihy_L'uboš"
    End If
    If ActiveSheet.Name = "Knihy_Žanetka" Then
        Worksheets("Knihy_Žanetka").Activate
        x = "Knihy_Žanetka"
    End If
    
    Set stlpce = ThisWorkbook.Worksheets(x).Range("$B$3:$J$3500")   'select columns of authors
    Set stlpce2 = ThisWorkbook.Worksheets(x).Range("$L$3:$L$3500")  'and ilustrator
    
    If stlpce.EntireColumn.Hidden = False Then  'if they are not hidden,
        stlpce.EntireColumn.Hidden = True       'hide them and
        stlpce2.EntireColumn.Hidden = True
        Worksheets(x).Buttons("Button 5").Caption = "Show more role"    'change label of button
    ElseIf stlpce.EntireColumn.Hidden = True Then   'if they are hidden...
        stlpce.EntireColumn.Hidden = False
        stlpce2.EntireColumn.Hidden = False
        Worksheets(x).Buttons("Button 5").Caption = "Hide more role"
    End If
    
End Sub
