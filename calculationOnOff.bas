Sub TurnOffCalc()
'to run script faster
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
End Sub

Sub TurnOnCalc()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
