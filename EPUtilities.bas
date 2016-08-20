Attribute VB_Name = "EPUtilities"
Option Explicit

Sub Relabel()

    '
    ' Procedure to relabel role table headers in the excel file
    ' This can be used to assist in the relabeling of table headers
    ' in the excel file.
    '
    
    Dim roles As Object
    Set roles = CreateObject("Scripting.Dictionary")
    Dim i As Range
    Dim key As Variant
      
    ' Creating a big dictionary of old and new labels
    
    ' Add roles mapping
    roles.Add "curr_tch_schadmin", "Curriculum/Teaching/School Administration"
    roles.Add "data_analysis", "Data Analysis"
    roles.Add "dev_grant_making", "Development/Grant Making"
    roles.Add "finance_budget_acc", "Finance/Budgeting/Accounting"
    roles.Add "human_rsrcs_capital_talent", "Human Resources/Human Capital/Talent"
    roles.Add "in_school", "In school"
    roles.Add "information_tech", "Information Technology"
    roles.Add "law_legal_services", "Law and Legal Services"
    roles.Add "marketing_comms_sales", "Marketing/Communications/Sales"
    roles.Add "no_info_available", "No information available"
    roles.Add "operations", "Operations"
    roles.Add "policy_advocacy_rsrch", "policy_advocacy_rsrch"
    roles.Add "General/Project Management", "project_management"
    roles.Add "Strategic Planning", "strategic_planning"
    roles.Add "unemployed", "Unemployed"
    roles.Add "working_role_unknown", "Working, role unknown"
    
    ' Add location mappings

    roles.Add "bay_area", "Bay Area"
    roles.Add "boston", "Boston"
    roles.Add "chicago", "Chicago"
    roles.Add "dc_metro", "DC Metro"
    roles.Add "denver", "Denver"
    roles.Add "new_orleans", "New Orleans"
    roles.Add "los_angeles", "Los Angeles"
    roles.Add "new_york", "Tri-state Area"
    roles.Add "tennessee", "Tennessee"
    roles.Add "texas", "Texas"
    
    ' Add Ethnicity mappings
    roles.Add "american_indian_alaskan_native", "American Indian, Alaskan Native"
    roles.Add "asian", "Asian"
    roles.Add "black_african_american", "Black African American"
    roles.Add "hispanic_latino", "Hispanic, Latino"
    roles.Add "multiethnic", "Multiethnic"
    roles.Add "pacific_islander", "Pacific Islander"
    roles.Add "white", "White"
    
    ' Add degree mappings
    roles.Add "bachelor", "Bachelor"
    roles.Add "certificate", "Certificate"
    roles.Add "doctor", "Doctor"
    roles.Add "master", "Master"
    
    ' Add fellowship type mapping
    roles.Add "gsf_10_week", "GSF 10 Week"
    roles.Add "gsf_yearlong", "GSF Yearlong"
    roles.Add "part_time_placement", "Part-time Placement"
    roles.Add "summer_track", "Summer Track"
    roles.Add "visiting_fellow", "Visiting Fellow"
    
    ' Add General category mappings
    roles.Add "other", "Other"
    roles.Add "total", "Total"
    
    
    
    For Each i In Selection
        For Each key In roles.Keys
            Select Case i.Value
                Case key
                    i.Value = roles(key)
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
Sub ConvertPercent()
    On Error GoTo handler
    Dim numRange, denomRange As Range
    Dim i As Variant
    Set numRange = Application.InputBox(prompt:="Select row with numerator values.", Type:=8)
    Set denomRange = Application.InputBox(prompt:="Select row with total(denominator)", Type:=8)
    

    Call GetPercent(numRange, denomRange)

handler:
    Exit Sub
    
End Sub

Sub GetPercentHelper(startCol As String, endCol As String, totalCol As String, _
    loopStartIndex As Integer, loopEndIndex As Integer)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets("role_fellowship_type")
    
    For i = loopStartIndex To loopEndIndex
        Call GetPercent(ws.Range(startCol & i & ":" & endCol & i), ws.Range(totalCol & i))
    Next i
    
End Sub

Sub epChart(sheetName As String, ByRef dataRange As Range, axisFontSize, gapWidth As Integer)
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
        .Parent.Width = 792
        .Axes(xlValue).HasMajorGridlines = False
        .ApplyDataLabels
        With .Axes(xlValue).TickLabels.Font
            .Size = axisFontSize
            .Bold = True
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

Sub placeEPChart()
    
    ' Use this procedure to write code to get percents. Run the code from this
    ' procedure

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim chartRange As Range
    Set chartRange = Application.InputBox(prompt:="Select data range to graph.", Type:=8)
 

    Call EPUtilities.epChart(Application.ActiveSheet.Name, chartRange, 14, 50)

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

Sub LabelValues()
    Dim labelRange As Range
    Dim seriesNum As Integer
    seriesNum = Application.InputBox("Enter series num", Type:=1)
    Set labelRange = Application.InputBox("Select the range that holds label values", Type:=8)
    Call assignDataLabelValues(seriesNum, labelRange)
    
End Sub


Sub AssignKeyShortcuts()
    Application.OnKey "^+5", "ConvertPercent"
    Application.OnKey "^+r", "Relabel"
    Application.OnKey "^+g", "placeEPChart"
    Application.OnKey "^+d", "LabelValues"
    
End Sub

