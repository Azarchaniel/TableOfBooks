Sub zorad_nazov_autor()
'
' Update - creating histogram, graphs and sorting by name and author
'
Dim table As String
Dim ws As String

If ActiveSheet.Name = "Knihy_L'uboš" Then
    Worksheets("Knihy_L'uboš").Activate
    ws = "Knihy_L'uboš"
    table = "Tabu1"
End If
If ActiveSheet.Name = "Knihy_Žanetka" Then
    Worksheets("Knihy_Žanetka").Activate
    ws = "Knihy_Žanetka"
    table = "Tabu2"
End If
If ActiveSheet.Name = "LP" Then
    Worksheets("LP").Activate
    table = "Tabu3"
End If
If ActiveSheet.Name = "Èasopisy" Then
    Worksheets("Èasopisy").Activate
    table = "Tabu4"
End If
Range("A1").Select

'I dont want it to create graphs in LP or magazines sheet
If ActiveSheet.Name = "Knihy_L'uboš" Or ActiveSheet.Name = "Knihy_Žanetka" Then
    Worksheets(ws).Range("$AG$16:$AP$36").Clear 'prepare area for graphs
    Worksheets(ws).Range("$AG$16:$AP$36").Delete
    Call histogramVysky 'calling functions
    Call Graf
    Range("$AF:$AF").Delete 'I couldnt identify a reason, why a new column is created everytime
    'functions are running. Therefore I'm just deleting column
    Columns("$AF:$AF").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("$AF:$AF").ClearFormats
End If


ActiveSheet.ListObjects(table).Sort. _
    SortFields.Clear                    'sort by title
ActiveSheet.ListObjects(table).Sort. _
    SortFields.Add Key:=Range(table + "[[#All],[Názov]]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
With ActiveSheet.ListObjects(table).Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
ActiveSheet.ListObjects(table).Sort. _
    SortFields.Clear                    'sort by author
'magazines doesn't have an author
If ActiveSheet.Name = "Knihy_L'uboš" Or ActiveSheet.Name = "Knihy_Žanetka" Or ActiveSheet.Name = "LP" Then
    ActiveSheet.ListObjects(table).Sort. _
        SortFields.Add Key:=Range(table + "[[#All],[Autor]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
End If
With ActiveSheet.ListObjects(table).Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
ActiveSheet.Calculate

Range("A1").Select
End Sub
