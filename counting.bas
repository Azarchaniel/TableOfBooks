
Option Explicit
'
'Function for counting amount of rows that are count-striked (means that they are no longer in my possesion)
'
Function CountStrikeThrough(myRng As Range) As Long
    Application.Volatile
    
    Dim myCell As Range
    Dim ctr As Long
    
    ctr = 0
    For Each myCell In myRng.Cells
        If myCell.Value = "" Then   'ignore empty cells
        Else
            If myCell.Font.Strikethrough = True Then
                ctr = ctr + 1
            End If
        End If
    Next myCell
    
    CountStrikeThrough = ctr 'return count
    
End Function
