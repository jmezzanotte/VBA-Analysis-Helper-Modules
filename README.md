# VBA-Analysis-Helper-Modules

#Author
John Mezzanotte

#Created
1/3/2016

# Overview
Visual Basic for Applications modules that have assisted my in day-to-day analysis tasks. Some of these modules may be used for general 
use and other modules are very custom to specific ananlysis tasks I have had. 

#Package Contents 
formatSignificantResults.bas


#formatSignificantResults 

#Description 
I used this module to help locate and format figures with significant results after multivariate tests have been ran on data. 

#Usage

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
