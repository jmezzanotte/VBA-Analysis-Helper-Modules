# VBA-Analysis-Helper-Modules

#Author
John Mezzanotte

#Created
1/3/2016

# Overview
Visual Basic for Applications modules that have assisted me in day-to-day analysis tasks. Most of these modules can be downloaded and used for general purpose analysis and graphing tasks within excel.

#Package Contents 
- formatSignificantResults.bas
- GraphUtilities.bas
- GraohUtilitiesMain.bas

#formatSignificantResults 

#Description 
I used this module to help locate and format figures with significant results after multivariate tests have been ran on data. 

#Specifications 
Sub FormatSignificantResult(sheetName As String, startRow As Integer, numRows As Integer, pvalColIndex As Integer, diffColIndex As _     Integer)
- Formats a cell value containing a significant result as XX%* (for example 10%*)
- Param - sheetName : The name of the sheet to be processed as a String 
- Param - startRow : The row number of where the data starts as Integer
- Param - numRows : The number of rows contained in the spreadsheet (or number of rows wished to process). 
- Param - pvalColIndex : The index of the column containing the p-value as Integer 
- Param - offset : The offset from the position of the p-val column to the cell that you wish to format as Integer
- Precondition: The significant figure is based on proportions 
- Postcondition: Cells with significant results are formatted to XX%*
- Return: None

#Usage
This module can be customized to suit your specific dataset. 

```

' pvalCol - Used to index the column of the spreadsheet that has  the p-value
' In this loop the first p-value column is located at column 8. 
' There are 38 columns in this particular spreadsheet 
' the p-value column is located every 5 columns (using Step 5, to loop to those columns) 
' The column that we want to format with the significant results in located 2 columns 
' before the p-value column (so we use pvalCol - 2, to dynamically find that cell).

Sub main()

    Dim pvalCol As Integer
    
    For pvalCol = 8 To 38 Step 5
        'running a check of the loop
        Call FormatSignificantResult("Sheet1", 2, 45, pvalCol, pvalCol - 2)
        
    Next pvalCol
    

End Sub

```
#GraphUtilities

#Description 
I used this module heavily on a project that required an extensive amount of custom graphing. This module is fairly customized
for to the project I created for, however it can be used when working with any excel dataset. It creates a custom stacked 100% stacked column chart with a data table at the foot of the graph. In the future I would like to make this module more customizable by expanding the API to allow for more options on the graph. 

#Procedures in Module
#Relabel 
- Description :This procedure relabels cell headers in a given table. The procedure takes a sigle argument which is basically a hashtable that maps raw column values to the new column labels. In VBA this procedure is expecting a dictionary object that maps old labels to new labels.

#GetPercent 

**Sub GetPercent(ByVal row As Range, ByVal total As Range)**

- **Description**: Calculates table percentages(either column or row based on need). Provide the function a range of cells to use as numerators and a second argument which is a cell to be used as a denominator for all the numerators provided in the first argument. Basically it will calcuate row or column percents based on the cell ranges passed to the function. It provides a method to calculate row and column percents in a table without having to manaully enter absolute cell ranges using in cell formulas.
- **Param row** :Row of valuse to use as a numerator 
- **Param total** : Single value to use a denominator for all values in the cell range speficied by row.

#epChart 
**Sub epChart(sheetName As String, ByRef dataRange As Range, axisFontSize, gapWidth As Integer, majorAxesUnit As Double, chartArea Outline As Integer, Optional chartHeight As Double = 4, Optional chartWidth As Double = 5, Optional chartTitle As String = "Title", Optional titleFontSize As Integer = 18)**
- **Description**:Places customized chart based on selected cell range( this chart was customized for the project, source code could be modified to suite your needs). 
- **Param sheetName**: Name of the sheet where the data is located and the location the graph will be placed.
- **Param dataRange**: Cell range of the data to be charted.
- **Param axisFontSize**: Font size of the graph axises.
- **Param gapWidth**: Width between column bars.
- **Param majorAxesUnit**: major axes value for the primary chart axes
- **Param chartAreaOutline**: Integer value for the type of chart outline to include. 
- **Optional Param chartHeight**: Double, height of chart, default = 4 
- **Optional Param chartWidth**: Double, width of chart, default = 5 
- **Optional Param graphTitle**: String, default = "Title"
- **Optional Param titleFontSize**:Integer, font size used for title, default = 18

# AssignDataLabelValues 

**Sub assignDataLabelValues(series As Integer, seriesRange As Range, Optional chartObjNum As Integer = 1)**

- **Description**: This function can be used to place custom data labels on graph bars. This function takes in a chart series number, as well as the range of values to apply to that series number display the custom labels. Use of this function replaces the need to hand map custom data labels to graph bars.
- **Param series**: Number of the data series you are applying labels to.
- **Param seriesRange**: The range of data that holds the data label values.
- **Param chartObjNum**: Integer that is the index number of the chart on a given sheet. Default value is 1, which will reference the  
                         first chart on a sheet. This parameter is used to target a specific chart when the procedure is being used 
                         on a sheet with multiple chart objects.
                         
# ChartObjectCount 

**Function ChartObjectCount()**
- **Description**: This function is used to count the number of chart objects on a sheet. This utility function is useful when used with the assignDataLabelValues procedure. Because the assignDataLabelValues procedure requires a index value of the chart being referenced it is import we always have an idea of how many charts are on a sheet. Using this function can help us to identify the correct chart for value labels when we have multiple charts. 
**return**: count as Long

# GraphUtilities Usage
I have included a module in this package called Main_demo.bas as a demonstration of how I have used these custom graph function from GraphUtilities in the past on projects. Below is the code from that file as well as some screen shots of what the data looks like before and after the functions have been applied to the data. 

# Use of the Relabel function 
In the Main_demo.bas file I wrote a function that calls the Relabel function from the GraphUtilities module. Below is a sample of the code I have used in the past to work with it. Below I create a dictionary object that maps old column and row labels to new labels and then I pass that object over to the Relabel function. I have mapped this function to a key stroke as well so it is easy to use while working in excel. 

```
Option Explicit

Sub DoRelabel()
    
    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")
    
    
    labels.Add "TestCol1", "Column 1"
    labels.Add "TestCol2", "Column 2"
    labels.Add "TestCol3", "Column 3"
    labels.Add "TestCol4", "Column 4"
    labels.Add "TestCol5", "Column 5"
    
    labels.Add "row_labels", ""
    
    labels.Add "variable1", "Demo Var 1"
    labels.Add "variable2", "Demo Var 2"
    labels.Add "variable3", "Demo Var 3"
    labels.Add "variable4", "Demo Var 4"
    labels.Add "variable5", "Demo Var 5"
    
    Call Relabel(labels)


```
#Before Relabel 
![relabel_raw](https://cloud.githubusercontent.com/assets/11713216/18944259/4c0b2ffa-85d9-11e6-9abc-b18ca1b9ece7.png)

# Calling DoRelabel
![relabel_after_proc](https://cloud.githubusercontent.com/assets/11713216/18944395/333d665e-85da-11e6-8669-a6b831910af2.png)


#Use of the GetPercent Function
In the Main_demo.bas file, I've also included a procedure for using the GetPercent function contained in the GraphUtilities module. This is helpful when you need to quickly calculate row are column percents in a table and don't want to manually enter cell formulas by hand with absolute denominators. Here in example of how I have used the GetPercent function with code taken from the Main_demo file.

```
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

```

#Before Running Convert Percent
![percent_raw](https://cloud.githubusercontent.com/assets/11713216/18976161/a8c4ca54-8664-11e6-8141-98e71c796a40.png)
![percent_raw_dialog](https://cloud.githubusercontent.com/assets/11713216/18976164/aec5121a-8664-11e6-850a-d685e656e542.png)
![percent_raw_dialog_2](https://cloud.githubusercontent.com/assets/11713216/18976169/b0e413e8-8664-11e6-80d2-69cfd465b472.png)

#After Running Convert Percent
![percent_converted](https://cloud.githubusercontent.com/assets/11713216/18976176/b926084a-8664-11e6-85a0-2a35fcbcc108.png)

#Use of epChart Function
In the Main_demo.bas file, I have also included a function that I have used in the past to use the epChart function. I have included an example that code below: 

```
Sub placeEPChart()
    
    ' Use this procedure to write code to get percents. Run the code from this
    ' procedure

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim chartRange As Range
    Set chartRange = Application.InputBox(prompt:="Select data range to graph.", Type:=8)
 
    
    Call EPUtilities.epChart(sheetName:=Application.ActiveSheet.Name, dataRange:=chartRange, _
        axisFontSize:=14, gapWidth:=50, majorAxesUnit:=0.2)

End Sub
```

#Before placeEPChart
![graph_function_select](https://cloud.githubusercontent.com/assets/11713216/19243479/b34ebc6c-8ecc-11e6-9779-b9e6aa8bdbd0.png)

#After placeEpChart
![graph_function_graph_placed](https://cloud.githubusercontent.com/assets/11713216/19243485/b6d6e31e-8ecc-11e6-8f84-9ec3eed453ff.png)

#Use of the assignDataLabelsValue function 
I have included code below called LabelValues that I have used to implement the assignDataLabelsValue procedure in the past. I typically 
assign these functions to keystrokes for easy use in excel. 

```
Sub LabelValues()
    On Error GoTo Handler
    
    Dim labelRange As Range
    Dim seriesNum As Integer
    seriesNum = Application.InputBox("Enter series num", Type:=1)
    Set labelRange = Application.InputBox("Select the range that holds label values", Type:=8)
    Call assignDataLabelValues(seriesNum, labelRange)

Handler:
    ' gracefully exit this procedure
    Exit Sub
    
End Sub
```
#Specify Series Number
![custom_data_labels](https://cloud.githubusercontent.com/assets/11713216/19245252/f13dd3ac-8ed4-11e6-97ba-52571ec36085.png)

#Specify Range of Values to Use
![custom_data_labels_range_select](https://cloud.githubusercontent.com/assets/11713216/19245257/f4969336-8ed4-11e6-8822-5c93dc6c4be0.png)

#Lables Applied
![custom_data_labels_applied](https://cloud.githubusercontent.com/assets/11713216/19245261/f7914324-8ed4-11e6-8131-9aaf63975f2d.png)
![custom_data_labels_fishished](https://cloud.githubusercontent.com/assets/11713216/19245270/fac7ad12-8ed4-11e6-97fe-2bb7d6fe1cb4.png)


#GraphUtilitiesMain.bas
This is a VBA module that demonstrates how I have used the procedures in the GraphUtilities module in my work. Feel to download this code and use it as a start point in your own work. 


# Collaborators 
John Mezzanotte 

# Ideas for Future Development 
In the formatSignificantResults module, allow the user to pass a format to the function. This way it is more customizable and you 
can use it in situations where you have figures other than proportions.
