# VBA-Analysis-Helper-Modules

#Author
John Mezzanotte

#Created
1/3/2016

# Overview
Visual Basic for Applications modules that have assisted me in day-to-day analysis tasks. Most of these modules can be downloaded and used for general purpose analysis and graphing tasks within excel.

#Package Contents 
-formatSignificantResults.bas
-GraphUtilities.bas


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
- Relabel : This procedure relabels cell headers in a given table. The procedure takes a sigle argument which is basically a hashtable that maps raw column values to the new column labels. In VBA this procedure is expecting a dictionary object that maps old labels to new labels.

- GetPercent : Calculates table percentages(either column or row based on need). Provide the function a range of cells to use as numerators and a second argument which is a cell to be used as a denominator for all the numerators provided in the first argument. Basically it will calcuate row or column percents based on the cell ranges passed to the function. It provides a method to calculate row and column percents in a table without having to manaully enter absolute cell ranges using in cell formulas. 

- epChart : Places customized chart based on selected cell range( this chart was customized for the project, source code could be modified to suite your needs). 

- AssignDataLabelValues : This function can be used to place custom data labels on graph bars. This function takes in a chart series number, as well as the range of values to apply to that series number display the custom labels. Use of this function replaces the need to hand map custom data labels to graph bars.

# GraphUtilities Usage
I have included a module in this package called Main_demo.bas as a demonstration of how I have used these custom graph function from GraphUtilities in the past on projects. Below is the code from that file as well as some screen shots of what the data looks like before and after the functions have been applied to the data. 

- Use of the Relabel function 
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

# Collaborators 
John Mezzanotte 

# Ideas for Future Development 
In the formatSignificantResults module, allow the user to pass a format to the function. This way it is more customizable and you 
can use it in situations where you have figures other than proportions.
