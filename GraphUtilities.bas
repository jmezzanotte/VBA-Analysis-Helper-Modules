Attribute VB_Name = "EPUtilities"
Option Explicit

Sub Relabel(labels As Object)

    '
    ' This procedure iterates through a selection as well as a scripting.dictionary.
    ' @params
    ' labels - As Object (This object will eventually be treated as a dictionary)
    ' Common Error - Type Error, usually occurs when you pass the method a type other than a dictionary
    '
    
    Dim i As Range
    Dim key As Variant
      
    For Each i In Selection
        For Each key In labels.Keys
            Select Case i.Value
                Case key
                    i.Value = labels(key)
            End Select
        Next key
    Next i


End Sub

Sub GetPercent(ByVal row As Range, ByVal total As Range)
    
    Dim i As Range
    
    For Each i In row
        i.Value = i.Value / total.Value
    Next i
    
    
End Sub

Sub epChart(sheetName As String, ByRef dataRange As Range, axisFontSize, gapWidth As Integer, majorAxesUnit As Double)
    ' Use embedded charts
    ' Embedded chart is in the chart object object, you can also use Shapes to refer to an embedded chart
    
    Dim epChart As Chart
    Dim seriesTotalIndex As Long
    Dim i, j As Long
    Dim totalSeries As series

    Set epChart = Charts.Add
    Set epChart = epChart.Location(Where:=xlLocationAsObject, Name:=sheetName)

    With epChart
        .ChartType = xlColumnStacked100
        ' Set data source range
        .SetSourceData Source:=dataRange, PlotBy:=xlColumns
        .GapDepth = 50
        seriesTotalIndex = .SeriesCollection.Count
        .ChartGroups(1).gapWidth = gapWidth
        .HasDataTable = True
        .DataTable.Font.Size = 11
        .HasLegend = False
        .Parent.Height = 432
        .Parent.width = 792
        With .Axes(xlValue, xlPrimary)
            .majorUnit = majorAxesUnit
            .HasMajorGridlines = False
            With .TickLabels.Font
                .Size = axisFontSize
                .Bold = True
            End With
        End With
    End With
   
    With epChart.SeriesCollection(seriesTotalIndex)
        
        .ChartType = xlLine
        .AxisGroup = xlSecondary
        .Format.Line.Visible = msoFalse
        .HasDataLabels = False
        
    End With
    
    With epChart.Axes(xlValue, xlSecondary)
        .TickLabelPosition = xlTickLabelPositionNone
        .Format.Line.Visible = False
    End With
    


End Sub

Sub assignDataLabelValues(series As Integer, seriesRange As Range)
    
    Dim epChart As Chart
    Dim i As Range
    Dim Counter As Integer
    
    Counter = 1
    ' Selects the first chart on the sheet.
    Set epChart = Application.ActiveSheet.ChartObjects(1).Chart
    epChart.SeriesCollection(series).HasDataLabels = True
   
    For Each i In seriesRange
        With epChart.SeriesCollection(series)
             Debug.Print .Name
            With .Points(Counter).DataLabel
                .Text = Round(i.Value * 100, 0) & "%"
                If .Text = "0%" Or .Text = "1%" Or .Text = "2%" Or .Text = "3%" Then
                    .Delete
                End If
            End With
            With .DataLabels.Font
                .Size = 12
                .Bold = True
            End With
            
            
        
        End With
        
        Counter = Counter + 1
    Next i
   
End Sub

Sub AdjustColumnWidths(colRange As Range, width As Integer)
    
    Dim i As Range
    
    For Each i In colRange
        i.ColumnWidth = width
    Next i
    

End Sub



