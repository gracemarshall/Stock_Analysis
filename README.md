# VBA Challenge All Stock Market Analysis
### Purpose
The VBA Challenge is to assist a client, Steve in decision making on what stocks to invest in. A data set of the Stock Market is provided with a code that needs to be refactored to measure performance. Refactoring is when you improve an existing code for a data set by reducing the number of steps required such as to loop, using less memory, and improved logic once to maximizing the run time for improved performance.

### Refactor VBA Code Analysis & Output Data 
Below is the VBA stock market code Analysis and output data for 2017 and 2018.  
- VBA data worksheet XLSM `VBA_Challenge.xlms` file for the project.

`VBA_Challenge.xlms`

![`VBA_Challenge.xlms`](https://github.com/gracemarshall/Stock_Analysis/blob/main/Resource/VBA_Challenge.xlsm)

- Run-time pop-up PNG messages screenshots after running refactored analyses for   2017 and 2018.

`VBA_Challenge_2017.png`

![VBA_Challenge_2017.png](https://github.com/gracemarshall/Stock_Analysis/blob/main/Resource/VBA_Challenge_2017.PNG)

`VBA_Challenge_2018.png`

![VBA_Challenge_2018.png](https://github.com/gracemarshall/Stock_Analysis/blob/main/Resource/VBA_Challenge_2018.PNG)


- 2017 and 2018 PNG stock output screenshots after running refactored analyses     for comparison.
`2017 Stock Output`

![2017 Stock Output.png](https://github.com/gracemarshall/Stock_Analysis/blob/main/Resource/2017%20Stock%20Output.PNG)

`2018 Stock Output`

![2017 Stock Output.png](https://github.com/gracemarshall/Stock_Analysis/blob/main/Resource/2018%20Stock%20Output.PNG)


- VBA_Challenge.vbs script to the Microsoft Visual Basic editor. The steps     Refactor VBA code and measure performance to add code where indicated by the   numbered comments in the starter code file.

        '1a) Create a tickerIndex and set to zero before looping over all the rows
        tickerIndex = 0
        
        '1b) create three output arrays and define data type
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
        For k = 2 To RowCount
            
        '3a)'increase totalVolume
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(k, 8).Value
                            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(k, 1).Value = tickers(tickerIndex) And Cells(k - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(k, 6).Value
            
        End If
                
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(k, 1).Value = tickers(tickerIndex) And Cells(k + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(k, 6).Value
        End If
        
        '3d Increase the tickerIndex.
        If Cells(k, 1).Value = tickers(tickerIndex) And Cells(k + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
        End If
    
        Next k
    
         '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
         For m = 0 To 11
         Worksheets("All Stocks Analysis").Activate
         Cells(4 + m, 1).Value = tickers(m)
         Cells(4 + m, 2).Value = tickerVolumes(m)
         Cells(4 + m, 3).Value = tickerEndingPrices(m) / tickerStartingPrices(m) - 1
                    
         Next m
    
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

## Results
Running the 2017 and 2018 stock data analysis provided stock performance and elapse time for both years. The refactored script for 2017 code ran for 0.1015625 seconds. The 2017 original script ran time was 3.414063 seconds. The refactored script for 2018 code ran for 0.1015625. The 2018 original script ran time was 2.734375E-02 seconds. 

The stock performance for 2017 outperformed the 2018 performance . The 2017 all stock analysis showed all positive stock gains with minimum gains of 5.5% for Run stock and maximum gains of 199.4% gain for DQ stock. Only one stock, Terp under performed with a -7.2%. The 2018 all stock analysis showed all negative stock loses with only two stocks gain, ENPH with 81.9% gain and  Run with 84% gain. The worse performing stocks are DQ at -62.6% and JKS at -60.5% with over 50% loss in in value.  

### Summary - Advantages and Disadvantages of Refactoring
**Advantages**
- Streamlined bolster and code upgrades. Clean code is much simpler to overhaul    and move forward and fast  for clients. 
- Refactoring can save maintenance cost as the support from developers will        require less time. 
- Code refactoring decreases the likelihood of errors  within the future and       simplifies the execution of program usefulness.
- Rather than making sense with tangling code or settling bugs, engineers can      begin actualizing the specified usefulness at once.
- Reduced complexity for easier understanding in the case of on-boarding a new     employee or team changes that may occur.  It is easier to refactor an existing   code than starting afresh. 

**Disadvantages**
- Refactoring can be costly and risky within the view of management. 
- Refactoring may present bugs. 
- Delivery plan is exceptionally tight. 
- testing outcomes can be affected by refactoring.
- Management doesn't care approximately viability and expansion of code base.

In conclusion, Code refactoring is an imperative step to evacuate code smells such as duplicate code, data clump, lack of design, long method, long parameter, divergent change and features. Code smell moderates down the improvement and is hence the need for a strong environment is needed to support code refactoring.

