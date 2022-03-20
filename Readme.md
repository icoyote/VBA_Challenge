# Stocks analysis with excel

## Overview of Project

Stock information varies by the minute and along the price of a unit, there is also how many (Volume of actions) are being traded. This kind of data is at best represented as line charts and extense excel tables. In this project, we have only the values of 12 distinct stocks identified by a code named ticker, the volume of traded units and price. The user of this Excel file has automated the calculations on volume and return based on Stock ticker and the yearly data that was saved as 2 Worksheets: "2018" and "2017".


### Purpose

The purpose of the prject is to optimize the execution time and allow more flexibility to the addition of more worksheets containing thousands of records.

The excel workbook VBA_Challenge.xlsm contains 2 spreadsheets named "2017" and "2018", each spreadsheet contains data for 11 stock tickers sorted by Date and by volume traded along with the closing price.
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

In order to avoid looping though thousands of rows every a a tickerIndex was suggested a function to retrive the tickerIndex was added to retrive the new index every time the ticker value changed.

This function returns the position of the ticker within it's declared Array tickers(), this position in the tickers() array will be used as the tickerIndex that will store the volume, starting price and ending price.

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

This function receives as arguments the value found for the ticker in the Worksheet "2017" or worksheet "2018" as well as the tickers() array that contains all 12 ticker codes. It position in the array was used as the tickerIndex.

Definition of arrays for results and tickerIndex
```
    'refactor step1b create 3 output Arrays
    Dim tickerVolumes() As Long
    Dim tickerStartingPrices() As Single
    Dim tickerEndingPrices() As Single
    
    'to measure execution time spent inside the loops
    startTime = Timer
```
```    
    tickerMaxSize = UBound(tickers)
    
    'refactor step1a creating a ticker index
    tickerIndex = 0
```

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

The refactored code executed in less than 0.08 seconds! 

In this very particular case the benefit in execution time is greatly worth the time spent degugging and making sure that the new code was not looping through the full amount of rows more than once.

The main advatage of refactoring code is to be able to reuse existing code and add new functionality or review the existing pattern of execution and optimize it. In this particular case when the original code the Stock Analysis was executed, the Excel application was frozen and none of the other applications that were in use were available while the analysis was taking place. To have the computer frozen for half a minute can be very inefficient and we don't know if the application failed or if the process is stuck in an infinite loop.

During this excercise, one aspect that wasn't as clear was the use of the active worksheet and some of the Debug functionalities that can be employed within the VB Developer screen. 

![image of new date columns](/resources/debug_ScreenShot.png)

revisiting and understanding the functionality was very useful. For the refactored version the worksheet containing the data to loop through was declared as a variable and the worksheet.Activate functionality was used for the resulting table in the worksheet that presented the Buttons and the VBA code.

```
    'Declare workbook object
    Dim wsht_year As Worksheet
    
    'retrive input year value
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Set the workbook that contains stock data by year
    Set wsht_year = ThisWorkbook.Worksheets("" + yearValue + "")
    
```

Use of the Activate() to set the focus in the Stocks Analysis worksheet

```
 'Output results to analysis sheet
   'activate the results spreadsheet
   
    'record results in output worksheet
    Worksheets("Stocks Analysis").Activate
    
    For r = LBound(tickers) To UBound(tickers)

       Cells(r + 4, 1).Value = tickers(r)
       Cells(r + 4, 2).Value = tickerVolumes(r)
       Cells(r + 4, 3).Value = (tickerEndingPrices(r) / tickerStartingPrices(r)) - 1
        
    Next r
    
    'activate the results spreadsheet
    Worksheets("Stocks Analysis").Activate
         
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
         
    'looping though values to highlight color of cell
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
```

Having a clear example by executing the original version of the code of the ACTUAL values expected (resulting table and resulting values for Volume and return) were essential in order to properly debug the code. The refactored code has to execute faster and provide exactly the same results already generated by the original script.

Refactoring not only means reusing code, it could also mean to introduce new modules of code or new functionalities to make the macros more effective at execution time or to perform validations that will minimize unexpected errors. For this example, if for whatever reason the sorting by ticker code and then by date (ascending date) in the 2017 and 2018 worksheets, the logic employed to increase the tickerIndex (checking if the ticker in the next row had changed) or if it was the first appearance of the ticker (to figure out the starting price) would fail.

Another danger would be to overcomplicate the refactored version and spend too much time re-coding the excel macros to obtain a more expensive (time spent by the developer and not executing faster and tieding up computer resources).


During the refactoring process and also during the debug process of the code provided in the module here are some of the reources used:
- [Arrays in vba](https://excelmacromastery.com/excel-vba-array/).
- [VBA Activate Worksheet in Excel](https://analysistabs.com/vba-code/worksheet/m/activate/).
