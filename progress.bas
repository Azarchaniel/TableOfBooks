Option Explicit

Sub InitProgressBar()

With progress
    .Bar.Width = 0
    .Text.Caption = "0% hotové"
    .Show vbModeless
End With

End Sub

Sub progressBar(CurrentProgress As Integer)

Dim ProgressPercentage As Double
Dim BarWidth As Long

BarWidth = progress.Border.Width * CurrentProgress * 0.01
ProgressPercentage = Round(CurrentProgress, 0)
    
progress.Bar.Width = BarWidth
progress.Text.Caption = ProgressPercentage & "% hotové"

DoEvents

End Sub
