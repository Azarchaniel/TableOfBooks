Sub Graf()
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
       answer = MsgBox("Chyba sa objavila príliš ve¾a krát. Spus debugger a stlaè F5.", vbOKOnly)
       Debug.Assert False
    End If
    
    Dim Pic As Object
    Dim ws As Worksheet

    If ActiveSheet.Name = "Knihy_L'uboš" Then
        Worksheets("Knihy_L'uboš").Activate
        ws = "Knihy_L'uboš"
    End If
    If ActiveSheet.Name = "Knihy_Žanetka" Then
        Worksheets("Knihy_Žanetka").Activate
        ws = "Knihy_Žanetka"
    End If
    Range("MiestoPreGraf").Select   'double check of deleting space where graphs will be
    Range("MiestoPreGraf").Delete
    Range("MiestoPreGraf").Clear

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
    
    Range("$AJ$16:$AP$25").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select 'create graph of type 2
    ActiveChart.SetSourceData Source:=Range("$AG$17:$AH$25")    'source data
    With ActiveChart.Parent
        .Top = Range("$AJ$16").Top 'set position of graph - upper border on this cell
        .Left = Range("$AJ$16").Left 'left border at this cell
        .Width = Range("$AJ$16:$AP$25").Width
        .Height = Range("$AJ$16:$AP$25").Height
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
    Range("$AJ$16").Select
    ActiveSheet.Pictures.Paste.Select   'paste pic of graph
    
    
    '
    ' /// Graph of width ///
    '
    ' same algortihm as above, but for different graph.
    'Maybe it could be written more general way and applied for both graphs at once,
    'but I doubt it would improve performance
    
    Range("$AJ$27:$AP$36").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("$AG$28:$AH$36")
    With ActiveChart.Parent
        .Top = Range("$AJ$27").Top
        .Left = Range("$AJ$27").Left
        .Width = Range("$AJ$27:$AP$36").Width
        .Height = Range("$AJ$27:$AP$36").Height
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
    Range("$AJ27").Select
    ActiveSheet.Pictures.Paste.Select
    
    Range("$AJ$37").Select
    
End Sub

