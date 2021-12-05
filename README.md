# VBA Challenge

## Project Overview

### Background
  Steve is happy witht he workbook that was created for him. He likes the features in his report.  With a snap of the finger, he can click a button and see his results.  He wants to expand his research for his parents to expand the dataset to the entire stock market, but thinks it will take too long to execute.  We'll take the current dataset and see if we can get the same results, but faster!

### Purpose

  To analyzea higher amount of stocks.  Steve is happy with the current results but let's see if we can't ge tthe results faster.  We'll take another look at the workbook and see where we can create even more effeciencies.  Here's what we'll do:
  - Look at the current code and data from the original code
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

![image](https://user-images.githubusercontent.com/94253815/144729980-be6d7b42-24c8-4f28-8e77-a83d046881d2.png)

![image](https://user-images.githubusercontent.com/94253815/144730052-1f224eaf-55d2-41ac-87e8-da04c8a1b204.png)


#### Refactoring the Code

In order to make the code more efficient I added the following:
 - Create (3) new arrays -  1) "Dim tickerVolume(12) As Long"- this will hold the volume. 2) " Dim tickerStartingPrices(12) As Single" -holds the starting price and 3)Dim tickerEndingPrices(12) As Single" - holds the end price.
 - We created a for loop to intialize the tickervolumes.
 - Inside the for loop, we created code that increases the stock ticker volume, which then adds the ticker volume for the current ticker.
 - Then we created two If / then statements on both the starting price and ending price to check on rows with the selected tickerIndex.  It wiill assign the current starting or ending price.
 - Finally, once the arrays are completed, we used for loops and variables to loop through the data, which will then compute our data and finish our analysis

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    YearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
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
    Worksheets(YearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    

    '1b) Create three output arrays
    Dim tickerVolume(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerIndex = i
        tickerVolumes(tickerIndex) = 0

    Next i

    ''2b) Loop over all the rows in the spreadsheet.
    tickerIndex = 0
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1) = tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
            If Cells(j + 1, 2).Value <> tickers(tickerIndex) And Cells(j, 1) = tickers(tickersIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
            End If

         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
             '3d Increase the tickerIndex.
                 tickerIndex = tickerIndex + 1
         End If
        'End If

    
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For k = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next k
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)

End Sub

### '17 vs. '18 Refactored Stock Performance
  When analyzing th stocks,it was a great investment decision in picking these stocks in 2017. 11 of 12 or 92% of the stockks had a positive return on investment.  Following up with thsame stocks in 2018, only 2 of 12 (17%) stocks had a positive return on investment.  It was a great decision in investing in RUN and ENPH, as these 2 stocks had positive outcomes in both 2017 and 2018. Based on these numbers, I would recommend looking into other stocks to invest in.
  
  
![image](https://user-images.githubusercontent.com/94253815/144730930-29d2535a-3240-4eaf-98d7-0da73bc95a9f.png)

![image](https://user-images.githubusercontent.com/94253815/144730958-5b39dc86-35d0-4ec6-a96a-217cdd3c5145.png)


### Coding Execution Time - Original vs. Refactored

After running the codes and tracking the times in both the original code and the refactored code, I saw minimal time difference.  I ran the code multiple times with different times.  These times captured represent thimes captured the most. 

Original Code Execution Time- 
 - 2017

![image](https://user-images.githubusercontent.com/94253815/144731054-3a3e988f-fc96-41bc-9338-ed53d736e188.png)


 - 2018

![image](https://user-images.githubusercontent.com/94253815/144731072-6f227418-7ab4-4519-8379-91d8d965ec7a.png)

Refactored Code Execution Time-

 - 2017

![image](https://user-images.githubusercontent.com/94253815/144731136-22ac741b-9314-44ba-8295-33ae61e23ae9.png)


 - 2018

![image](https://user-images.githubusercontent.com/94253815/144731150-99f2a1a5-7739-4599-bc7a-0d2bda6bb793.png)


## Summary


### Advantages of refactored Code


### Disadvantages of Refactored Code

### Pro's and Con's apply to refactoring the original VBA script?
