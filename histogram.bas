Sub histogramVysky()
    '
    'creating histogram of dimensions of books using Excel function
    '

    Dim i As Integer
    Dim ws As String
    
    'activate particular sheet to work with
    If ActiveSheet.Name = "Knihy_L'uboš" Then
        Worksheets("Knihy_L'uboš").Activate
        ws = "Knihy_L'uboš"
    End If
    If ActiveSheet.Name = "Knihy_Žanetka" Then
        Worksheets("Knihy_Žanetka").Activate
        ws = "Knihy_Žanetka"
    End If
    
    'format of cell is TEXT
    Range("$AG$16:$AG$36").NumberFormat = "@"
    
    '
    'Height range text
    '
    For i = 0 To 8
      Range("$AG$" & 17 + i).Value = (i * 5) & " - " & (i * 5 + 5)  'creating labels
      Range("$AH$" & 17 + i) = Application.WorksheetFunction.CountIfs(Worksheets(ws).Range("Výška"), "<=" & (i * 5 + 5), Worksheets(ws).Range("Výška"), ">" & (i * 5))
    Next i
    Range("$AG$25").Value = "<40"
    

    '
    'Width range text
    '
    For i = 0 To 8
      Range("$AG$" & 28 + i).Value = (i * 5) & " - " & (i * 5 + 5)
      Range("$AH$" & 28 + i) = Application.WorksheetFunction.CountIfs(Worksheets(ws).Range("Šírka"), "<=" & (i * 5 + 5), Worksheets(ws).Range("Šírka"), ">" & (i * 5))
      'count how many values are between borders
      'Im using named range
    Next i
    Range("$AG$36").Value = "<40"
    
    Range("$AG$16").Value = "Rozmer"
    Range("$AG$27").Value = "Rozmer"
    Range("$AH$16").Value = "Poèet kníh"
    Range("$AH$27").Value = "Poèet kníh"
    Range("$AG$16:$AH$16").Font.Italic = True
    Range("$AG$27:$AH$27").Font.Italic = True
    
    '
    'Drawing of borders for histogram values
    '
    Range("$AG$16:$AH$16").Select
    Range("$AG$16:$AH$16").HorizontalAlignment = xlLeft
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("$AG$25:$AH$25").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Range("$AG$27:$AH$27").Select
    Range("$AG$27:$AH$27").HorizontalAlignment = xlLeft
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("$AG$36:$AH$36").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    'set alignment
    Range("$AG$17:$AH$25").HorizontalAlignment = xlRight
    Range("$AG$28:$AH$36").HorizontalAlignment = xlRight
    Range("$AG$17:$AH$25").VerticalAlignment = xlBottom
    Range("$AG$28:$AH$36").VerticalAlignment = xlBottom

End Sub

