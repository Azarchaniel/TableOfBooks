Sub histogramVysky()
'
'creating histogram of dimensions of books using Excel function
'
    Dim i As Integer
    Range("AC15:AC35").NumberFormat = "@"
    Range("AC16:AC24").HorizontalAlignment = xlRight
    Range("AC27:AC35").HorizontalAlignment = xlRight

    '
    'HEIGHT
    '
    For i = 0 To 8
    Range("AC" & 16 + i).Value = (i * 5) & " - " & (i * 5 + 5) 'creating labels
      Range("AD" & 16 + i) = Application.WorksheetFunction.CountIfs(Range("V3:V1000"), "<=" & (i * 5 + 5), Range("V3:V1000"), ">" & (i * 5))
    'count how many values are between borders
    Next i
    Range("AC24").Value = "<40"

    '
    'WIDTH
    '
    For i = 0 To 8
      Range("AC" & 27 + i).Value = (i * 5) & " - " & (i * 5 + 5)
      Range("AD" & 27 + i) = Application.WorksheetFunction.CountIfs(Range("W3:W1000"), "<=" & (i * 5 + 5), Range("W3:W1000"), ">" & (i * 5))
    Next i
    Range("AC35").Value = "<40"
    
    Range("AC15").Value = "Dimension"
    Range("AC26").Value = "Dimension"
    Range("AD15").Value = "Amount of b."
    Range("AD26").Value = "Amount of b."

    '
    'drawing borders for histogram values
    '
    Range("AC15:AD15").Select
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("AC24:AD24").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Range("AC25:AD25").Select
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    
    Range("AC26:AD26").Select
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("AC35:AD35").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
End Sub
