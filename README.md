# stock-analysis

## Overview of Project

The purpose of this analysis was to compute the total daily volume and percentage of change annually of various stocks, in order to clearly visualize the percentage by which each stock increased or decreased in value over a given timeframe. The computation and visualization was performed for both 2017 and 2018 stock value figures. Visualization was improved by adding conditional formatting to highlight increased values in green and decreased values in red.



## Results

The overall result was that ENPH stock performed the best, and most stocks performed worse in 2018 than in 2017. This result was measured by dividing the closing stock price at the end of the timeframe of this study by the stock price at the beginning of the timeframe, to quantify the amount of change. The following screenshots demonstrate the amount of change in a table format.

![2017](https://user-images.githubusercontent.com/100863488/157741105-ec44e04d-1645-42b9-9c41-ade1a2ec075c.png)

![2018](https://user-images.githubusercontent.com/100863488/157741124-72656e2d-41a8-48ec-8852-790f38eaf4fa.png)


### Results of Refactoring

The analysis was improved, after its original coding, with refactored code. The refactored analysis utilized a `tickerIndex` variable across four arrays (`tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`).

The refactored code produced significant gains in terms of computation speed. Using the original code, the computation time for 2017 figures was approximately .746 seconds and for 2018 figures was approximately .754 seconds. Using the refactored code, the computation time was approximately .105 seconds for both 2017 and 2018. 

![VBA_Challenge_2017_originalcode](https://user-images.githubusercontent.com/100863488/157741205-b73e3aed-55e9-40df-9271-838b0f26209d.png)
![VBA_Challenge_2017_refactoredcode](https://user-images.githubusercontent.com/100863488/157741220-a4dcb804-596f-4691-b755-b38af6b6eaf0.png)

![VBA_Challenge_2018_originalcode](https://user-images.githubusercontent.com/100863488/157741229-2e0e1123-1407-406f-bdae-701c4d9092c6.png)
![VBA_Challenge_2018_refactoredcode](https://user-images.githubusercontent.com/100863488/157741244-fd56c5a7-3743-4480-9be8-237cc0799f26.png)


### Refactored Code

The entire refactored code follows: 

```
Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index and set to zero
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    '2a) Create a for loop to initialize the tickerVolumes to zero
    
    For tickerIndex = 0 To 11
    tickerVolumes(tickerIndex) = 0
    'tickerEndingPrices(tickerIndex) = 0
    'tickerStartingPrices(tickerIndex) = 0
    
    Sheets(yearValue).Activate
    
     
    '2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            End If
        
        '3c) Check if the current row is the last row with the selected ticker
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            End If
            
        '3d) Increase the tickerIndex.
            If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
            tickerIndex = tickerIndex + 1
            End If
             
          Next j
   
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
   
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Next i
    
    Next tickerIndex
  
   
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Font.Bold = True
    Range("A1").Font.Underline = xlUnderlineStyleSingle
    Range("A1").Font.Size = 16
    
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "The refactored All Stocks Analysis ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```



## Summary

### What are the advantages or disadvantages of refactoring code?
Refactoring involves editing or rewriting code in order to make it more streamlined, clearer, easier to process, and/or less error-prone, and it can have many advantages. In this example, refactoring the code led to significant gains in computation speed. Refactoring can also lead to code that is easier for multiple people to understand and collaborate on and code that makes a program more versatile for future use.

There are potential disadvantages to refactoring code, but they are outweighed by the advantages of doing so. The primary disadvantages would include the accidental introduction of errors or the misunderstanding of previous code such that the refactored code no longer serves the purpose of the original code. 


### How do these pros and cons apply to refactoring the original VBA script?
In this case, refactoring made our code more complex from the user perspective, but more nimble from the processor's perspective, as it used a a `tickerIndex` variable to link four arrays (`tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`). The program did not need to loop through the dataset using the refactored code in the same labor-intensive way as it did in the original code. This would be a great benefit if the code were used on a much larger dataset.

Since the code is longer and more complex in the refactored code, the refactored code is not necessarily more intuitive for the user, so commenting and the use of whitespace are more important in the refactored code than in the original.


