# Stock Analysis in VBA

## Overview
In this project, we originally wrote a VBA code to help Steve's family analyze a green energy companies stocks to see if it was worth investing in. Now for the challenge, we are refactoring that code to have it run more efficiently and to look at all 12 of the stocks over the past several years to see which performed the best. To do so, we created four ticker() arrays that looked at Tickers, TickerVolume, TickerStartingPrice, and TickerEndingPrice, with the difference of the last two giving us the return each stock yielded.

`
    
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

    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
            
        Next i
        
        '2b) Loop over all the rows in the spreadsheet.
            For i = 2 To RowCount
    
    
    '3a) Increase volume for current ticker
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
         
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
             tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        'End if
        
             '3c) Chen, C (2020) GitHub Stock Analysis [Source code]. https://github.com/caseychen3605/stock-analysis
              'Check if the current row is the last row with the selected ticker
              'If the next row’s ticker doesn’t match, increase the tickerIndex.
              'If  Then
            
                 If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                 tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            

                   '3d Increase the tickerIndex.
                    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                       tickerIndex = tickerIndex + 1
                
                    End If
           
                    'End If
    
    Next i
    
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
`
## Results
After running our code, we found that the stocks ENPH and RUN are the only two that had a postive return in both 2017 and 2018, and would be the two we would suggest Steve's family invested in out of the 12. Of those two though, ENPH was the highest returning and would be our number one pick, as highlighted in the images below. 

### 2017
<img width="758" alt="Stock Analysis 2017" src="https://user-images.githubusercontent.com/103979087/167310691-eaeb9830-1d5f-4331-be48-0356d8aaad98.png">


### 2018
<img width="758" alt="Stock Analysis 2018" src="https://user-images.githubusercontent.com/103979087/167310694-aa72bef3-d859-4812-b9b2-23840bdf0847.png">



## Summary

### Pros and Cons
Refactoring our code allowed for our macros to run quicker, and it cleaned up the code aesethically to make it easier to read and navigate through. This allows errors to be more easily found and fixed, and for others to understand our code easier. One con however would be the tedious nature of refactoring the original code. Renaming and filling in the new information can lead to easy slipups and misplacing of information.

Below are the elapsed times for our Macros, both of which ran faster than the original VBA code which was first running at about .09 and .08. As you can see, we cut that time down significantly and yielded more information.

### 2017
<img width="261" alt="VBA_Challenge2017" src="https://user-images.githubusercontent.com/103979087/167310902-5f67368c-c7c6-439f-b326-500aea7b4790.png">

### 2018
<img width="261" alt="VBA_Challenge2018" src="https://user-images.githubusercontent.com/103979087/167310911-789eba1b-9d10-4e27-9d5b-b0118678dfcf.png">

