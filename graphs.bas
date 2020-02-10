Sub graf()
    '
    'Function for creating graphs.
    '
    
    On Error GoTo ErrorCounter
    On Error Resume Next    'if error occures, usualy helps to run function second time
    
ErrorCounter:   'if error occures more than 5 times, user have to solve it in debugger
    Dim i As Integer
    Dim answer As String
    i = i + 1
    If i = 5 Then
    answer = MsgBox("Error occurs too many times. Run debuger and press F5.", vbOKOnly)
       Debug.Assert False
    End If
    
    Dim Pic As Object

    If ActiveSheet.Name = "Knihy_L'uboš" Then
        Worksheets("Knihy_L'uboš").Activate
    End If
    If ActiveSheet.Name = "Knihy_Žanetka" Then
        Worksheets("Knihy_Žanetka").Activate
    End If
    Range("MiestoPreGraf").Select   'double check of deleting space where graphs will be
    Range("MiestoPreGraf").Delete
    Range("MiestoPreGraf").Cleer

    For Each Pic In ActiveSheet.Pictures
        If Not Intersect(Pic.TopLeftCell, Range("$AB$15:$AN$35")) Is Nothing Then
            Pic.Delete
        End If
    Next Pic
    'find every picture in given range and delete it. Excel is not deleting pictures
    'with Delete or Clear functions
    
    '
    ' /// Graph of height ///
    '
    
    Range("$AK$16:$AQ$25").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select 'create graph of type 2
    ActiveChart.SetSourceData Source:=Range("$AH$17:$AI$25")    'source data
    With ActiveChart.Parent
        .Top = Range("$AK$16").Top 'set position of graph - upper border on this cell
        .Left = Range("$AK$16").Left 'left border at this cell
        .Width = Range("$AK$16:$AQ$25").Width
        .Height = Range("$AK$16:$AQ$25").Height
    End With
    ActiveChart.HasLegend = False
    With ActiveChart.Axes(xlCategory) 'add label to the lower axis
        .HasTitle = True
        .AxisTitle.Text = "Výška knihy v cm"
    End With
    With ActiveChart.Axes(xlValue) 'add label of left axis
        .HasTitle = True
        .AxisTitle.Text = "Poèet kníh"
    End With
    ActiveChart.Parent.Name = "Graf1"   'add graph a name so I can manipulate it
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Výška kníh"
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels 'activate graph's labels
    ActiveChart.ChartGroups(1).GapWidth = 52
    Application.CommandBars("Format Object").Visible = False
    ActiveChart.PlotArea.Select
    Application.CommandBars("Format Object").Visible = False
    
    ActiveSheet.ChartObjects("Graf1").Activate  'activate the specific graph
    ActiveChart.Parent.Cut                      'cut it to clipboard
    Range("$AS$16").Select                      'and paste it to different cell
    ActiveSheet.Paste
    'there was a problem with exporting graph as a picture, when it was only in clipboard
    
    ActiveSheet.ChartObjects("Graf1").Chart.CopyPicture xlScreen, xlBitmap  'select graph, copy it as a picture
    ActiveSheet.ChartObjects("Graf1").Delete    'delete graph
    Range("$AK$16").Select
    ActiveSheet.Pictures.Paste.Select   'paste pic of graph
    
    
    '
    ' /// Graph of width ///
    '
    ' same Algortihm as above, but for different graph.
    'Maybe it could be written more general way and applied for both graphs at once,
    'but I doubt it would improve performance
    
    Range("$AK$27:$AQ$36").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("$AH$28:$AI$36")
    With ActiveChart.Parent
        .Top = Range("$AK$27").Top
        .Left = Range("$AK$27").Left
        .Width = Range("$AK$27:$AQ$36").Width
        .Height = Range("$AK$27:$AQ$36").Height
    End With
    ActiveChart.HasLegend = False
    With ActiveChart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Šírka knihy v cm"
    End With
    With ActiveChart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Poèet kníh"
        
    End With
    ActiveChart.Parent.Name = "Graf2"
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Šírka kníh"
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.ChartGroups(1).GapWidth = 52
    Application.CommandBars("Format Object").Visible = False
    ActiveChart.PlotArea.Select
    Application.CommandBars("Format Object").Visible = False
    
    ActiveSheet.ChartObjects("Graf2").Activate
    ActiveChart.Parent.Cut
    Range("$AS$27").Select
    ActiveSheet.Paste
    
    ActiveSheet.ChartObjects("Graf2").Chart.CopyPicture xlScreen, xlBitmap
    ActiveSheet.ChartObjects("Graf2").Delete
    Range("$AK27").Select
    ActiveSheet.Pictures.Paste.Select
    
    Range("$AK$37").Select
    
End Sub
