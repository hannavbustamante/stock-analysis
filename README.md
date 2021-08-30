# Overview of Project
For the Module 2 challenge, I used my knowledge on VBA to refactor the Stocks Analysis code from Module 2. This will allow the code to run for much larger datasets without having to worry about the code taking a long time to execute.

# Results
## Original Code
```
Sub AllStocksAnalysis():

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
    
    Worksheets("All Stocks Analysis").Activate
    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
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
    
    Dim startPrice As Single
    Dim endPrice As Single
    
    
    Sheets(yearValue).Activate
    
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        Sheets(yearValue).Activate
        
        For j = 2 To rowEnd
            If Cells(j, 1).Value = ticker Then
            'finding total volume for current ticker
                totalVolume = Cells(j, 8).Value + totalVolume
            End If
    
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                'set starting price
                startingPrice = Cells(j, 6).Value
            End If
     
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'set ending price
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j
           
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
    
    endTime = Timer
    MsgBox ("This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue))

End Sub

Sub formatAllStocksAnalysisTable():
    
    Worksheets("All Stocks Analysis").Activate
    'Formatting
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Color = vbBlue
    Range("A3:C3").Font.Size = 14
    Range("A3:C3").Font.TintAndShade = -0.5
    
    
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3).Value > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
            ' colors the cell green
   
        
        ElseIf Cells(i, 3).Value < 0 Then
            Cells(i, 3).Interior.Color = vbRed
            'colors the cell red
        
        Else
            Cells(i, 3).Interior.Color = xlNone
            'clears the cell color
        
        End If
    
    Next i

End Sub

```
### Runtime for original code

![Original code Runtime 2017](https://user-images.githubusercontent.com/88729583/131275186-d54e9d1f-c0bf-471b-b72a-0ebe0de06064.PNG)
![Original code Runtime 2018](https://user-images.githubusercontent.com/88729583/131275188-15cb4e78-7510-46d1-ac50-5ee2818d3188.PNG)

## Refactored code
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
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0

    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
    
         If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
                'set starting price
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
         End If
  
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                'set ending price
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
             End If
    
    Next j
    Next tickerIndex

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

### Runtime for Refactored code

![Refactored code Runtime 2017](https://user-images.githubusercontent.com/88729583/131275329-e4f3a0b3-dabe-40f9-a06f-7b0420eadd9e.PNG)
![Refactored code Runtime 2018](https://user-images.githubusercontent.com/88729583/131275330-9a199572-78b0-4305-8b38-ce287b58789b.PNG)

## Efficiency improvement

For the year 2017, the runtime went from ~0.859 seconds to ~0.156 seconds. The runtime is ~82% faster.

For the year 2018, the runtime went from ~0.867 to seconds to 0.164 seconds. The runtime is ~81% faster.


# Summary
## Advantages and disadvantages of refactoring code
### Advantages
The advantages of refactoring code are that it improves the design of your code, makes your program run faster and run for larger sets of data.
The disadvantages of refactoring code are that it can be time consuming, there might not be a more efficient way to write the code so it could be a waste of time, and it could introduce bugs in your code causing it to not run properly or at all.

### Advantages and disadvantages of the original and refactored VBA script
These advantages and disadvantages apply to the VBA script I refactored. The code now runs around 80% faster which is a huge improvement and won't take a long time to run for massive data sets. It is also easier to read the code and the design has been refined. However, refactoring the code did take a lot of time so in a situation where there is a deadline or a time constraint, it might have been better to keep the original code. In addition, when refactoring the code I ran into a lot of new bugs in the code that I had to fix.
