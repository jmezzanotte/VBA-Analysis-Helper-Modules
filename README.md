# VBA-Analysis-Helper-Modules

#Author
John Mezzanotte

#Created
1/3/2016

# Overview
Visual Basic for Applications modules that have assisted me in day-to-day analysis tasks. Some of these modules may be used for general purposes and other modules are very custom to specific ananlysis tasks I have had. 

#Package Contents 
formatSignificantResults.bas
EpUtilities.bas


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
#EPUtilities

#Description 
I used this module heavily on a project that required an extensive amount of custom graphing. This module is fairly customized
for to the project I created for, however it can be used when working with any excel dataset. In the future (time permitting) i would like to make this module more robust. 

#Procedures in Module
- Relabel : Taylored specifically to project tables. It relabels cell headers in the table. Source code could be modified to include or change out labeling values. 
- ConvertPercent : Calculates table percentages(either column or row based on need) 
- epChart : Places customized chart based on selected cell range( this chart was customized for the project, source code could be modified to suite your needs). 
- LabelValues : used to apply custom value labels to bars. 

# Collaborators 
John Mezzanotte 

# Ideas for Future Development 
In the formatSignificantResults module, allow the user to pass a format to the function. This way it is more customizable and you 
can use it in situations where you have figures other than proportions.
