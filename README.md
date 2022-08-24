# VBA_Challenge
Stock analysis with VBA

# Overview of Project #

### Purpose ###
The purpose of this project was to refactor a VBA code to gather stock information for the years 2017-2018 and to determine whether or not the stocks are good investsments.

## Data ##
The data shown includes two charts with stock performances for the years 2017-2018 on 12 different stocks. The charts give you the **Ticker, Total Daily Volume,** and the yearly **Return** from the stocks.


# Analysis #
Before I had a chance to refactor the code I had to download it and open it using a different app because it wasn't allowing me to open it with Excel. Once I was finally able to access the code I copied and pasted it to my VBA page. I would go back and forth to my old code I had used before during the module to use as refrence and would have to make a few adjustments so the code would run smoothly for this particular assingnment. This was challenging on some parts due to me overlooking some small variable changes. Below I have included the code that I used to come up with the stock analysis with a few comments to guide me and the reader on what is going on as the code is being written out.




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
    Dim tickervolumes(12) As Long
    Dim tickerstartingprices(12) As Single
    Dim tickerendingprices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickervolumes(i) = 0
        
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickervolumes(tickerIndex) = tickervolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerstartingprices(tickerIndex) = Cells(i, 6).Value
        End If
            
    
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerendingprices(tickerIndex) = Cells(i, 6).Value
            
        End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickervolumes(i)
        Cells(4 + i, 3).Value = tickerendingprices(i) / tickerstartingprices(i) - 1
        
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
    
    <img width="1123" alt="Screen Shot 2022-08-24 at 4 27 52 PM" src="https://user-images.githubusercontent.com/110702997/186527450-2c080321-f9b1-4094-8d8f-7654db1306ea.png">




