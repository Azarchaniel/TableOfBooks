Sub Graf()
    '
    'Function for creating graphs. It's kinda slow, but it's because Excel is creating graphs very slow
    'and there is a lot settings applying
    '
    
    Dim Pic As Object

    If ActiveSheet.Name = "Knihy_L'uboš" Then   'if name of sheet is X, activate X
        Worksheets("Knihy_L'uboš").Activate
    End If
    If ActiveSheet.Name = "Knihy_Žanetka" Then
        Worksheets("Knihy_Žanetka").Activate
    End If
    Range("AE15:AN35").Select   'double check of deleting space where graphs will be
    Range("AE15:AN35").Delete
    Range("AE15:AN35").Clear

    For Each Pic In ActiveSheet.Pictures
        If Not Intersect(Pic.TopLeftCell, Range("AB15:AN35")) Is Nothing Then
            Pic.Delete
        End If
    Next Pic
    'find every picture in given range and delete it. Excel is not deleting pictures
    'with Delete or Clear functions
    
    '
    ' /// Graph of height ///
    '
    
    Range("AC15:AD24").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select 'create graph of type 2
    ActiveChart.SetSourceData Source:=Range("$AC$16:$AD$24")    'source data
    With ActiveChart.Parent
        .Top = Range("AF15").Top 'set position of graph - upper border on this cell
        .Left = Range("AF15").Left 'left border at this cell
        .Width = Range("AF15:AL24").Width
        .Height = Range("AF15:AL24").Height
    End With
    With ActiveChart.Axes(xlCategory) 'add label to the lower axis
        .HasTitle = True
        .AxisTitle.Text = "Height of b. in cm"
    End With
    With ActiveChart.Axes(xlValue) 'add label of left axis
        .HasTitle = True
        .AxisTitle.Text = "Poèet kníh"
    End With
    ActiveChart.Parent.Name = "Graf1"   'add graph a name so I can manipulate it
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Height of b."
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels 'activate graph's labels
    ActiveChart.ChartGroups(1).GapWidth = 52
    Application.CommandBars("Format Object").Visible = False
    'Sheets("List1").Select
    ActiveChart.PlotArea.Select
    Application.CommandBars("Format Object").Visible = False
    
    ActiveSheet.ChartObjects("Graf1").Activate  'activate the specific graph
    ActiveChart.Parent.Cut  'cut it to clipboard
    Range("AS15").Select    'and paste it to different
    ActiveSheet.Paste
    'there was a problem with exporting graph as a picture, when it was only in clipboard
    
    ActiveSheet.ChartObjects("Graf1").Chart.CopyPicture xlScreen, xlBitmap  'select graph, copy it as a picture
    Range("AE15:BB35").Select   'prepare area for pasting
    Range("AE15:BB35").Delete
    Range("AE15:BB35").Clear
    Range("AF15").Select
    ActiveSheet.Pictures.Paste.Select   'normal paste doesnt work for pictures
    
    
    '
    ' /// Graph of width ///
    ' same algortihm as above, but for different graph.
    'Maybe it could be written more general way and applied for both graphs at once,
    'but I doubt it would improve performance
    
    Range("AC26:AD35").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("$AC$27:$AD$35")
    With ActiveChart.Parent
        .Top = Range("AF26").Top
        .Left = Range("AF26").Left
        .Width = Range("AF26:AL35").Width
        .Height = Range("AF26:AL35").Height
    End With
    With ActiveChart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Width of b. in cm"
    End With
    With ActiveChart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Amount of b."
        
    End With
    ActiveChart.Parent.Name = "Graf2"
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Width of b."
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.ChartGroups(1).GapWidth = 52
    Application.CommandBars("Format Object").Visible = False
    'Sheets("List1").Select
    ActiveChart.PlotArea.Select
    Application.CommandBars("Format Object").Visible = False
    
    ActiveSheet.ChartObjects("Graf2").Activate
    ActiveChart.Parent.Cut
    Range("AE26").Select
    ActiveSheet.Paste
    
    ActiveSheet.ChartObjects("Graf2").Chart.CopyPicture xlScreen, xlBitmap
    Range("AE26:BB35").Select
    Range("AE26:BB35").Delete
    Range("AE26:BB35").Clear
    Range("AF26").Select
    ActiveSheet.Pictures.Paste.Select
    
    
    Range("AE37").Select
    
    
End Sub
