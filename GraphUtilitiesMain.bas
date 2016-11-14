Attribute VB_Name = "Main"
Option Explicit

Sub DoRelabel()
    
    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")
    
    
    labels.Add "strongly_agree", "Strongly agree"
    labels.Add "agree", "Agree"
    labels.Add "somewhat_agree", "Somewhat agree"
    labels.Add "disagree", "Disagree"
    labels.Add "strongly_disagree", "Strongly disagree"
    labels.Add "very_good", "Very good"
    labels.Add "good", "Good"
    labels.Add "poor", "Poor"
    labels.Add "average", "Average"
    labels.Add "very_poor", "Very poor"
    
    labels.Add "All", "All sites"
    labels.Add "missing", "Missing"

   
    
    labels.Add "much_too_fast", "Much too fast"
    labels.Add "slightly_too_fast", "Slightly too fast"
    labels.Add "just_right", "Just right"
    labels.Add "slightly_too_slow", "Slightly too slow"
    
    Call Relabel(labels)

End Sub
Sub placeEPChart()
    
    ' Use this procedure to write code to get percents. Run the code from this
    ' procedure
    On Error GoTo Handler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim chartRange As Range
    Dim chartTitle As String
    Set chartRange = Application.InputBox(prompt:="Select data range to graph.", Type:=8)
    chartTitle = InputBox(prompt:="Enter chart title:")


    Call EPUtilities.epChart(sheetName:=Application.ActiveSheet.Name, dataRange:=chartRange, _
        axisFontSize:=14, gapWidth:=30, majorAxesUnit:=0.2, chartAreaOutline:=xlNone, _
        chartHeight:=4.2, chartWidth:=9.5, graphTitle:=chartTitle, titleFontSize:=14)
    
Handler:
    Exit Sub

End Sub

Sub LabelValues()
    On Error GoTo Handler
    
    Dim labelRange As Range
    Dim seriesNum As Integer
    Dim chartIndex As Integer
    chartIndex = Application.InputBox("Enter the chart index (this is useful if there is more than one chart object on a page." & _
    "The first chart object is used by default")
    seriesNum = Application.InputBox("Enter series num", Type:=1)
   
    
    Set labelRange = Application.InputBox("Select the range that holds label values", Type:=8)
    
    Call assignDataLabelValues(seriesNum, labelRange, chartIndex)

Handler:
    ' gracefully exit this procedure
    Exit Sub
    
End Sub

Sub InitAdjustColumnWidths()
    
    On Error GoTo Handler
    
    Dim colRange As Range
    Dim width As Integer
    Set colRange = Application.InputBox("Enter column range", Type:=8)
    width = Application.InputBox("Enter desired column width", Type:=1)
    
    Call AdjustColumnWidths(colRange, width)
    
    
Handler:
    Exit Sub
    
End Sub

Sub ConvertPercent()
    
    On Error GoTo Handler
    
    Dim numRange, denomRange As Range
    Dim i As Variant
    Set numRange = Application.InputBox(prompt:="Select row with numerator values.", Type:=8)
    Set denomRange = Application.InputBox(prompt:="Select row with total(denominator)", Type:=8)
    

    Call GetPercent(numRange, denomRange)

Handler:
    Exit Sub
    
End Sub


Sub AssignKeyShortcuts()
    Application.OnKey "^+5", "ConvertPercent"
    Application.OnKey "^+r", "DoRelabel"
    Application.OnKey "^+g", "placeEPChart"
    Application.OnKey "^+d", "LabelValues"
    Application.OnKey "^+-", "InitAdjustColumnWidths"
    
End Sub
