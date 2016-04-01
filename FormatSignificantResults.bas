Attribute VB_Name = "FormatSignificantResults"

Option Explicit

Sub FormatSignificantResult(sheetName As String, startRow As Integer, numRows As Integer, pvalColIndex As Integer, _
    offset As Integer)

    Dim wrkbk As Workbook
    Dim wrksht As Worksheet
    Dim targetCell As Range
    Dim i As Integer

    Set wrkbk = ActiveWorkbook
    Set wrksht = wrkbk.Worksheets(sheetName)

    For i = startRow To numRows
        If wrksht.Cells(i, pvalColIndex).Value <= 0.05 Then
            wrksht.Cells(i, offset).NumberFormat = "0.0%\*"
        ElseIf wrksht.Cells(i, offset).Value <> "." Then
             wrksht.Cells(i, offset).NumberFormat = "0.0%"
        End If
    Next i
    
End Sub



