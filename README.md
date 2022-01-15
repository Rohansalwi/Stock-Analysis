# **Stock-Analysis**
## **1. Overview of Project**
This project is intended to extract information from the data dump using VBA to help Steve find the total daily volume and yearly return for each stock for the years 2017 and 2018 and make a decision.
## **1. Results**
For the year 2017, It is found that the stock ENPH had the maximum return of 129.5% and for the year 2018 stock "RUN" has the maximum return of 84%.

<img width="600" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/96554223/149631939-36f01ae9-9be7-4feb-a9cc-a919f4d13a4d.png">
<img width="607" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/96554223/149631930-8f5e8f95-2099-4e65-adba-818710b1b31c.png">

The snippet of code is as follows:

 
 Sub AllStocksAnalysisRefactored()
  
    Dim startTime As Single
    Dim endTime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
     tickerIndex = 0


    '1b) Create three output arrays
    
    Dim totalVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
       For i = 0 To 11
       totalVolumes(i) = 0
       Next i
        
    ''2b) Loop over all the rows in the spreadsheet.

    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
               totalVolumes(tickerIndex) = totalVolumes(tickerIndex) + Cells(i, 8).Value
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
             If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

           End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            
        tickerIndex = tickerIndex + 1
        End If
        
        Next i
 
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
            
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = totalVolumes(i)
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub

Stock ENPH is the second most returning stock just little behind RUN at 81.9%. Rest of the stocks are in loss for the year 2018 as compared to 2017.
## **1. Summary**
Original code may be time consuming as the logic has to be coded in VBA if there is not one already. The advantage of refactoring a code is that since the logic is already there it saves time and leads to better quality code. The disadvantage is that you need to undertand the original code before refactoring it and it can be thus time consuming and be iterative. Also, there is a lot of testing.

