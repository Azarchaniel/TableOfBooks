Sub applyStyle()
'
' ApplyStyle Makro
'
'
Dim wRng1 As Range 'working range
Dim wRng2 As Range
Dim tRng As Range
Dim rng As Range
Dim row As Range
Dim offset As Integer
'every sheet is slightly different, so I have to adjust settings

If ActiveSheet.Name = "Knihy_L'uboš" Then
    Worksheets("Knihy_L'uboš").Activate
    Set wRng1 = Range("A4:AF4")
    Set wRng2 = Range("A5:AF5")
    Set tRng = Range("A4:A2500")
    offset = 31
End If
If ActiveSheet.Name = "Knihy_Žanetka" Then
    Worksheets("Knihy_Žanetka").Activate
    Set wRng1 = Range("A4:AF4")
    Set wRng2 = Range("A5:AF5")
    Set tRng = Range("A4:A2500")
    offset = 31
End If
If ActiveSheet.Name = "LP" Then
    Worksheets("LP").Activate
    Set wRng1 = Range("A4:L4")
    Set wRng2 = Range("A5:L5")
    Set tRng = Range("A4:A500")
    offset = 11
End If
If ActiveSheet.Name = "Èasopisy" Then
    Worksheets("Èasopisy").Activate
    Set wRng1 = Range("A4:H4")
    Set wRng2 = Range("A5:H5")
    Set tRng = Range("A4:A500")
    offset = 7
End If

'startTime = Now() 'for checking performance speed
Call TurnOffCalc
wRng1.Select
Call Grey   'apply Grey style
wRng1.Copy  'copy whole row
Call progressBar(41)    'set progress bar to 41 perc.
For Each row In tRng.Rows
    row.Select
    If ActiveCell.row Mod 2 = 0 Then
      Range(row, row.offset(0, offset)).PasteSpecial Paste:=xlPasteFormats  'for every second row, paste formats
    End If
Next row
Call progressBar(70)
wRng2.Select
Call White
Set rng = Range("A4:A2500")
wRng2.Copy
For Each row In tRng.Rows
    row.Select
    If ActiveCell.row Mod 2 = 1 Then
      Range(row, row.offset(0, offset)).PasteSpecial Paste:=xlPasteFormats
    End If
Next row
Call progressBar(98)
Call TurnOnCalc
'finTime = Now()    'for checking performance speed
'resultTime = DateDiff("s", startTime, finTime)
'MsgBox resultTime

End Sub

Sub Grey()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    ActiveWindow.ScrollColumn = 2
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    End Sub
    
    Sub White()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149998474074526
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub
