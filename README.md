# Stock_analysis

## Overview of Project
Our client Steve valued the original stock analysis tool.  Now he would like to expand the scope of the analysis to a larger number of stocks.  To allow the stock analysis tool to handle a larger number of stocks, the code will be refactored to improve the performance of the tool by reducing its run time.



## Results

### Stock Investment Results
The overall portfolio performance of the selected stocks over 2018 showed a significant drop when compared to the performance of 2017.  See the below annual volume and return summary for both years below.  This reflects the overall market (using the DJIA as a reference); there was a large selloff at the end of 2018.  However two stocks, ENPH and RUN, still showed sizable gains.  A deeper analysis of the fundamentals of these firms to determine if additional investment is warranted.


* [2017 Portfolio Returns](https://github.com/goldbala55/stock_analysis/blob/main/resources/VBA_Challenge_2017_stock_returns.png)

* [2018 Portfolio Returns](https://github.com/goldbala55/stock_analysis/blob/main/resources/VBA_Challenge_2018_stock_returns.png)




### Refactoring Results
    The refactoring effort was highly successful. Sample run times using the current data samples (12 stocks, ~3000 rows/year) decreased from 0.456 seconds to 0.180 seconds, a drop of 60%.  Images of the run times are available below. 


* [Sample run time pre refactoring](https://github.com/goldbala55/stock_analysis/blob/main/resources/VBA_Challenge_2017_preFactor.png)

* [Sample run time post refactoring](https://github.com/goldbala55/stock_analysis/blob/main/resources/VBA_Challenge_2017_postFactor.png)

The major contributor to these results was eliminating the nested loop.  

    The original approach required a complete rereading of the entire input table for each stock to be analyzed. 

Original pseudo code
<details>
  <summary>Click to expand</summary>
'Loop through tickers 
    For i = 0 To 11
        'initialize
        
        'Active data sheet
        Worksheets(yearValue).Activate
        'now process the entire data sheet for each year
        For j = 2 To rowEnd
            If Cells(j, 1).Value = ticker Then
                'Bump the total if the stock (col 1) is the current ticker
                totalVolume = totalVolume + Cells(j, 8).Value
            
                'Grab the starting price
                If Cells(j - 1, 1).Value <> ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
            
                'Grab the last price
                If Cells(j + 1, 1).Value <> ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
            End If
        
        Next j
            
        'Activate reporting worksheet
        Worksheets("All Stock Analysis").Activate
        
        'write the results. one stock per row
        
        'Conditionally format results
        '    
    Next i
</details>

    The refactored code only requires a single pass of the input table.  This is accomplished by using an index and a set of arrays to hold interim results.
Refactored pseudo code
<details>
  <summary>Click to expand</summary>
'Define and Initialize array of all tickers

'Define the index, and required output arrays

Dim tickerIndex As Integer

Dim tickers(12) As String

Dim tickerVolumes(12) As Long

Dim tickerStartingPrices(12) As Single

Dim tickerEndingPrices(12) As Single

'Initialize the tickerVolumes to zero.
        
'Activate data worksheet
Worksheets(yearValue).Activate

'Step through the rows in the spreadsheet:
'  1. total the volume for each stock
'  2. extract the starting and ending prices

tickerIndex = 0
For i = 2 To RowCount
        
    'Bump the total if the stock (col 1) is the current ticker
    If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    End If
        
    'Grab the starting price
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
        
    'Grab the last price
    'We are at the last row for a given symbol, so we need bump the ticker index too
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        tickerIndex = tickerIndex + 1
    End If
    
Next i

'Point to the Reporting Sheet 
'Loop through the arrays to output the Ticker, Total Daily Volume, and Return.
For tickerIndex = 0 To 11
    
    Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
    Cells(4 + tickerIndex, 2).Value = tickerVolumes(tickerIndex)
    Cells(4 + tickerIndex, 3).Value = (tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1
    
Next tickerIndex

'Style formatting for the results
</details>


## Summary 

1. Refactoring can bring about substantial improvements in processing time (as in this case) as well as add new functionality. However, refactoring requires an investment of person-hours that may not always yield results corresponding to the effort and can also introduce new bugs into the code.
2. The original code provided all the functionality requested by the client and given the modest size of the tables, performed acceptably. However this code will not scale well. Assuming ~250 rows/year per stock, a 100 stock table will have 25,000 rows.
   1.   Analyzing all 100 stocks with the original code will require processing 2.5 million rows.
   2.   Using the refactored code will only require processing 25,000 rows.  A substantial reduction. 