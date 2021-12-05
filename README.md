# VBA Challenge

## Project Overview

### Background
  Steve is happy witht he workbook that was created for him. He likes the features in his report.  With a snap of the finger, he can click a button and see his results.  He wants to expand his research for his parents to expand the dataset to the entire stock market, but thinks it will take too long to execute.  We'll take the current dataset and see if we can get the same results, but faster!

### Purpose

  To analyzea higher amount of stocks.  Steve is happy with the current results but let's see if we can't ge tthe results faster.  We'll take another look at the workbook and see where we can create even more effeciencies.  Here's what we'll do:
  - Look at the current code and data from the original code
  - Show original code
  - Provide the data from the original code
  - Show refactored code
  - 2017 / 2018 refactored stock performance
  - Coding execution time - Original vs. Refactored
    
## Results


### Original Code

Sub AllStocksAnalysis()

Dim startTime As Single
Dim endTime As Single

Worksheets("All stocks analysis").Activate

YearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
'1) Format the output sheet on All Stocks Analysis worksheet
    'Worksheets("All Stocks Analysis").Activate
    'Range("A1").Value = "All Stocks (2018)"
    'Create a header row

    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
 '---------------------------------------------------
'2)Initialize array of all tickers

    Dim tickers(11) As String
    
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
    
'3)Initialize variables for starting price and ending price
   'Dim startingPrice As Single
   'Dim endingPrice As Single

    Dim startingPrice As Single
    Dim endingPrice As Single

'3b) Activate data worksheet
   'Worksheets("2018").Activate
    
   Worksheets(YearValue).Activate
    
'3c) Get the number of rows to loop over
   'RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through tickers
    'For i = 0 To 11
    'ticker = tickers(i)
    'totalVolume = 0
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
    
   
    
'5) loop through rows in the data
    'For j = 2 To RowCount
    
    Worksheets(YearValue).Activate
    
        For j = 2 To RowCount
'5a) Get total volume for current ticker
    If Cells(j, 1).Value = ticker Then
    
        totalVolume = totalVolume + Cells(j, 8).Value
    
     End If
     
'5b) get starting price for current ticker
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        startingPrice = Cells(j, 6).Value
    
    End If
    
'5c) get ending price for current ticker

    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        endingPrice = Cells(j, 6).Value
    
    End If

Next j
   
   
'6) Output data for current ticker

    Sheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
 'Code is good through 2.4.1--------------------------------------------------
 
Next i

    endTime = Timer
    MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year " & (YearValue)
    

'Formatting

Worksheets("All Stocks Analysis").Activate

Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A3:C3").Borders(xlEdgeTop).LineStyle = xlContinuous


Range("A3:C3").Font.FontStyle = "Bold"
Range("A3:C3").Font.ColorIndex = 3
Range("B4:B15").NumberFormat = "$#,##0.00"
Range("C4:C15").NumberFormat = "0.00%"
Columns("A:C").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
End Sub

#### Original Code Data

#### Refactor Code

### '17 vs. '18 Refactored Stock Performance


### Coding Execution Time - Original vs. Refactored


## Summary


### Advantages of refactored Code


### Disadvantages of Refactored Code

### Pro's and Con's apply to refactoring the original VBA script?
