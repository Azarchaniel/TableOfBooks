Option Explicit

Sub InitProgressBar()

With Progress
    .Bar.Width = 0
    .Text.Caption = "0% hotové"
    .Show vbModeless
End With

End Sub

Sub progressBar(CurrentProgress As Integer)

Dim ProgressPercentage As Double
Dim BarWidth As Long

Progress.Bar.Width = Progress.Border.Width * CurrentProgress * 0.01
Progress.Text.Caption = Round(CurrentProgress, 0) & "% hotové"

DoEvents

End Sub
