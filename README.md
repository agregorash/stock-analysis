# Green Stocks Analysis

## Overview of Project

A financial advisor, Steve, has asked me to help him analyze stock performance in the green energy sector on behalf of his parents.  His parents are keen in investing in a particular green energy stock, DAQO.  I used the Visual Basic Application in Excel to create a macro that judged stock performance on two criteria; daily trade volume and annual return for the years of 2017 and 2018.  I then created a macro to analyze a group of 11 more stocks in addition to DAQO, in order to give them a context of this stock's performance so they could compare results and determine if a more diversified investment strategy would be appropriate.  The final product delivered to Steve included refactored code which reduced run time and computing power used.        

### Purpose

The purpose of this project was to create a more efficient way to analyze a large data set in Excel using VBA.  By creating macros in VBA I was able to provide accurate analysis of a large data set, of stock performance in the green energy sector, for two separate years at the click of a button.  I was then able to create an even more efficient VBA script, reducing run time and computing power, by refactoring the original code.

## Results

### Stock Performance and Run Times

The analysis show that green stocks performed much better in 2017 than 2018.  In 2017 all but one stock analyzed generated a positive return, the majority reporting sizeable returns in excess of 20%.  Stock performance in 2018 was much different.  All but two stocks analyzed generated a negative return, almost half of which reported sizeable losses on returns in excess of 20%.  The two stocks that did generate postitive return did perform very well with returns of 82% and 84%

The original code used generated results in about .84 seconds for the 2017 analysis and .85 seconds for the 2018 analysis.  In order to make the code more efficient I changed the nesting order of the loops by creating 4 different arrays; tickers, tickerVolumes, tickerStartingPrices and tickerEndingPrices.  The 'tickers' array established the symbol of the stock and was matched with the three other arrays using the variable'tickerIndex', allowing me to assign the other variable to each individual ticker symbol before iterating through the data set.  This resulted in the refactored code generating the same results in about .15 and .13 seconds respectively, significantly reducing run time and increasing efficiency. 

Find the difference in code and run times listed below

#### Original Code and Performance

    Dim startTime As Single
    Dim endTime As Single

        yearValue = InputBox("What year would you like to run the analysis on?")
        
            startTime = Timer
    
    '1) Format the output sheet on All Stocks Analysis worksheet
   
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2) Initialize array of all tickers
   
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
        
        '3a) Initialize variables for starting price and ending price
        
        Dim staringPrice As Single
        Dim endingPrice As Single
        
        
    '3b) Activate data worksheet
   
    Worksheets(yearValue).Activate
   
    '3c) Get the number of rows to loop over
   
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
    '4) Loop through tickers
   
    For i = 0 To 11
       ticker = tickers(i)
       TotalVolume = 0
       
       '5) loop through rows in the data
       
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
           
            If Cells(j, 1).Value = ticker Then
        
                TotalVolume = TotalVolume + Cells(j, 8).Value
            
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
       
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = TotalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    
    Next i
   
   ![2017 Original](https://github.com/agregorash/stock-analysis/blob/main/Original%20Run%20Times/Original%202017.PNG)
   ![2018 Original](https://github.com/agregorash/stock-analysis/blob/main/Original%20Run%20Times/Original%202018.PNG)
   
#### Refactored Code and Performance

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
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
    For i = 0 To 11
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
            
        'End If
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
        
    Next i

![2017 Refactored](https://github.com/agregorash/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
![2018 Refactored](https://github.com/agregorash/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

## Summary of Results

### Refactoring Code

The main advantage of refactoring code is that it makes the code more efficient.  The amount of computing power and run time can be sigificantly reduced, especially when using very large data sets.  The main disadvantage of refactoring code is that you are risking making your working code unusable if you do not save the original code.  Always be sure to save your code!

### Refactoring VBA Script 

Refactoring the VBA script resulted in a more efficient macro that ran in a fraction of the time the original code, and used less computing power.  These differences may not seem like a big deal in this case, but in a much larger data set the reductions could be crucial.  The process of refactoring code is also an efficient process, as you are able to compare your old and new code side by side.  On the other hand, the risk of losing your code is a prevalent disadvantage in VBA if you are not saving your work.  Another disadvantage of refactoring in VBA is that if you do not have a good understanding of the syntax and theory, you may struggle to create a working macro during the debugging process. 
