Sub zorad_nazov_autor()
'
' Update - creating histogram, graphs and sorting by name and author
'I'm sorting firstly by name and then by author. Sorting only by author
'
Dim x As String

If ActiveSheet.Name = "Knihy_L'uboš" Then
    Worksheets("Knihy_L'uboš").Activate
    x = "Tabu1"
End If
If ActiveSheet.Name = "Knihy_Žanetka" Then
    Worksheets("Knihy_Žanetka").Activate
    x = "Tabu2"
End If
If ActiveSheet.Name = "LP" Then
    Worksheets("LP").Activate
    x = "Tabu3"
End If
Range("A1").Select

        'I dont want for it to create graphs in LP sheet
If ActiveSheet.Name = "Knihy_L'uboš" Or ActiveSheet.Name = "Knihy_Žanetka" Then
    Range("AC14:AO36").Clear    'prepare area for graphs
    Range("AC14:AO36").Delete
    Call histogramVysky     'calling functions
    Call Graf
    Range("AB16:AB23").Clear
    Range("AB:AB").Delete   'I couldnt identify a reason, why a new column is created everytime
    'functions are running. Therefore I'm just deleting column
    Columns("AB:AB").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AB1:AB2").Delete
End If


ActiveSheet.ListObjects(x).Sort. _
    SortFields.Clear    'sort by title
ActiveSheet.ListObjects(x).Sort. _
    SortFields.Add Key:=Range(x + "[[#All],[Title]]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
With ActiveSheet.ListObjects(x).Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
ActiveSheet.ListObjects(x).Sort. _
    SortFields.Clear 'sort by author
ActiveSheet.ListObjects(x).Sort. _
    SortFields.Add Key:=Range(x + "[[#All],[Author]]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
With ActiveSheet.ListObjects(x).Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
ActiveSheet.Calculate
Range("N1").Select
End Sub
