# Stocks analysis with excel

## Overview of Project

Stock information varies by the minute and among the price of a unit is also how many are being traded. This kind of data is at best represented as line charts and extremely populated tables. In this project, we have only the values of 12 distinct stocks identified by a code named ticker, the volume of traded units and price.


### Purpose

The excel workbook contains 2 spreadsheets, each spreadsheet contains data for 11 stock tickers sorted by Date and by volume traded along with the closing price.
A summary worksheet has been created it has been called "Stocks Analysis" and it's content has been populated using VBA macros. These macros have been assigned to 2 buttons.

- Clear Worksheet
- Run Analysis

![image of new date columns](/resources/Buttons_Screen.png)

In order to have a empty worksheet our user can hot the button Clear Worksheet

```
Sub ClearWorksheet()
    Cells.ClearFormats
    Cells.Clear
End Sub

```

An input box appears when the "Run Analysis" button is activated by the user, the year to run the analysis is requested. in order to function a worksheet named as the year requested must exist and have valid data.

![image of new date columns](/resources/Input_Screen.png)


The worksheets populated with data for the 2017 and 2018 years have thousands of rows, to summarize each ticker individually and even done multiple times would be too time consuming if we have more than 11 tickers only. Macros and VBA code will be useful to automate such operations in order to allow larger number of tickers to be added and more worksheets to increase the year analysis to decades or to refine it by months or days.

VBA code was written to loop through all the tickers one by one and to retrieve the total daily volume and return of each ticker by year. 

Before refactoring the stock analysis code the execution time neared the 35seconds.

Execution time before refactoring of code for 2017 and 2018 data:

![image of new date columns](/resources/originalExecTime2017.png)
![image of new date columns](/resources/originalExecTime2018.png)


In the original version of the Sub AllStocksAnalysis() we had 12 tickers stored in an array and  looping 12 time through the complete worksheet

![image of new date columns](/resources/analysis2017_screen.png)

```

    'Initializing Array to hold the 12 tickers codes
    Dim tickers(12) As String
    
    'hardcoding of the tickers
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
```

In order to avoid looping though thousands of rows every a a tickerIndex was suggested and instead if looping multiple times through the whole worksheet, a function to retrive the tickerIndex was added to retrive the new index every time the ticker value changed.

code of function that retrieves the position of the ticker code being analysed, this position in the tickers() array will be used as the tickerIndex that will store the Volume, starting price and ending price

Code of the function to retrive the next Index value:
```

Function IsInArray(stringToBeIndexed As String, arr As Variant) As Long
    Dim i As Long
    'default return value if nothing found
    IsInArray = -1
    For i = LBound(arr) To UBound(arr)
        If StrComp(stringToBeIndexed, arr(i), vbTextCompare) = 0 Then
            IsInArray = i
            Exit For
        End If
    Next i
End Function

```

This function receives as arguments the value found for the ticker in the Worksheet "2017" or worksheet "2018" as well as the tickers() array that contains all 12 ticker codes. It position in the array was used as the tickerIndex

Use of the tickerIndex within the full worksheet loop to populate Volume, starting price and enfing price

```
 'refactor step2b For loop over all the rows in the spreadsheet
    
    For j = rowStart To rowEnd
    
        'refactor step 3a
        If j = 2 Then
            'first row of the sheet and the first ticker symbol we will retrive from the worksheet
            'get the tickerIndex
            tickerIndex = IsInArray(wsht_year.Cells(j, 1).Value, tickers)
        End If
        
               
        If wsht_year.Cells(j, 1).Value = tickers(tickerIndex) Then
                  'increase tickerVolume for this tickerIndex
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + wsht_year.Cells(j, 8).Value
                
        End If
            
        If wsht_year.Cells(j - 1, 1).Value <> tickers(tickerIndex) And wsht_year.Cells(j, 1).Value = tickers(tickerIndex) Then
                
                'this is the first occurence of stock tickers(tickerIndex)
                'the row above has a ticker OTHER than tickers(tickerIndex)
                tickerStartingPrices(tickerIndex) = wsht_year.Cells(j, 6).Value
    
            End If
    
            If wsht_year.Cells(j + 1, 1).Value <> tickers(tickerIndex) And wsht_year.Cells(j, 1).Value = tickers(tickerIndex) Then
                'this is the last occurence of stock tickers(tickerIndex)
                'the following row has a ticker OTHER THAN tickers(tickerIndex)
                tickerEndingPrices(tickerIndex) = wsht_year.Cells(j, 6).Value
                
                'retrieve the next ticketIndex when the next row's ticker doesn't match the current row
                tickerIndex = IsInArray(wsht_year.Cells(j + 1, 1).Value, tickers)
                
    
            End If
            
        Next j
```

## Results:

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

| Results by Original Script                                         | Execution Time |
| ------------------------------------------------------------------ | -------------- |
| ![image of new date columns](/resources/originalExecTime2017.png)  | 33.64062 secs  |
| ![image of new date columns](/resources/originalExecTime2018.png)  | 33.67578 secs  |

| Results by Refactored Script                                       | Execution Time |
| ------------------------------------------------------------------ | -------------- |
| ![image of new date columns](/resources/VBA_Challenge_2017.png)    | 0.0703125 secs |
| ![image of new date columns](/resources/VBA_Challenge_2018.png)    | 0.078125 secs  |

The refactored code executed in less than 0.08 seconds! In this very particular case the benefit in execution time is greatly worth the time spent degugging and making sure that the new code was not looping through the full amount of rows more than once.



What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

During the refactoring process and also during the debug process of the code provided in the module here are some of the reources used:
- [Arrays in vba](https://excelmacromastery.com/excel-vba-array/).
- [VBA Activate Worksheet in Excel](https://analysistabs.com/vba-code/worksheet/m/activate/).