# stock-analysis

##VBA analysis for Steve

##PURPOSE

The purpose of this project is to refactor a Microsoft Excel VBA code to collect a variety of stock information from the year 2017 and the 
year 2018, and determine whether or not the stocks are worth investing. This process has been originally completed in a similar format, however, 
the goal for this attempt is to increase the efficiency of the original code.

###The Data
The data that is presented includes two charts with stock information on 12 different stocks. The stock information does have the date the stock 
was issued, the ticker value, the opening, close and adjusted closing price, the highest and lowest price, and the volume of the stock. 
The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

##RESULTS

###Analysis
While refactoring the code, I started by adding the code that was required to activate the appropriate worksheet. The steps are listed 
out in order to set the structure for the refactoring. Here are the instructions with the code as written in the file:

'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output array
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
     
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    
   For i = 2 To RowCount
   
    
        '3a) Increase volume for current ticker
        
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        
        '3c) check if the current row is the last row with the selected ticker
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If                

        '3d Increase the tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
        tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

##SUMMARY

###The Advantages & Disadvantages of Refactoring Stock Analysis
As a result of the refactoring, it is a decrease in macro run time. which is favorable comparing it to the original analysis, whereas our
new analysis took less time to run. Attached is a Resource folder with screenshots that indicate the run time for our new analysis.

###Pros and Cons of Refactoring Code
The process of Refactoring helps make our code more organized. Some of the advantages from a cleaner code include software and design
improvement, faster programming , and debugging. It becomes easier to read. 
Disadvantages from Refactoring is related to the eextencion for the application that we are inquiring.

