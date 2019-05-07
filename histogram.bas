Sub histogramVysky()
'
'creating histogram of dimensions of books using Excel function
'
        Dim i As Integer

    '
    'HEIGHT
    '
    For i = 0 To 8
    Range("AC" & 16 + i).Value = (i * 5) & " - " & (i * 5 + 5) 'creating labels
      Range("AD" & 16 + i) = Application.WorksheetFunction.CountIfs(Range("V3:V1000"), "<=" & (i * 5 + 5), Range("V3:V1000"), ">" & (i * 5))
    'count how many values are between borders
    Next i
    Range("AC24").Value = "<40"
    Range("AC15:AC35").NumberFormat = "@"
    Range("AC16:AC24").HorizontalAlignment = xlRight
    Range("AC27:AC35").HorizontalAlignment = xlRight

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
End Sub
