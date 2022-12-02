##
**Stock Analysis Overview**
#
**Purpose of Analysis**

The purpose of this analysis was to help Steve analyze a higher number of stocks throughout 2017 and 2018, also to do a little more research for his parents to better support their investments.  
#
**Results**

To help Steve with the deeper dive of the analysis I did some refactoring with the sample code provided. With this code I created new loops, ticker array to worksheet and values. Below is the image of the refactored code:
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
    
    Dim tickerIndex As Integer
    'Initiate tickerIndex at zero.
    tickerIndex = 0
    

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create for loop to initialize the tickerVolumes to zero.
    
    For tickerIndex = 0 To 11
    'Initiate each ticker's volume at zero.
    tickerVolumes(tickerIndex) = 0
        
    '2b) Loop over all the rows in the spreadsheet.
    
        For i = 2 To RowCount
        
    '3a) Increase volume for current ticker
    
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
    '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
                    
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
           End If
            
            
    '3c) Check if the current row is the last row with the selected ticker.
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
           'End if
           End If
            
      '3d) Increase the tickerIndex.
      
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
           
              tickerIndex = tickerIndex + 1
           
        'End If
        End If
    
        Next i
        
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
#
**2017 vs 2018 Stock Performance** 

All Stocks (2017)		
		
Ticker	Total Daily Volume	Return
AY	136,070,900	8.9%
CSIQ	310,592,800	33.1%
DQ	35,796,200	199.4%
ENPH	221,772,100	129.5%
FSLR	684,181,400	101.3%
HASI	80,949,300	25.8%
JKS	191,632,200	53.9%
RUN	267,681,300	5.5%
SEDG	206,885,200	184.5%
SPWR	782,187,000	23.1%
TERP	139,402,800	-7.2%
VSLR	109,487,900	50.0%
![image](https://user-images.githubusercontent.com/118132063/205223946-7a345549-9918-476b-b075-74b4c956b07d.png)
All Stocks (2018)		
		
Ticker	Total Daily Volume	Return
AY	83,079,900	-7.3%
CSIQ	200,879,900	-16.3%
DQ	107,873,900	-62.6%
ENPH	607,473,500	81.9%
FSLR	478,113,900	-39.7%
HASI	104,340,600	-20.7%
JKS	158,309,000	-60.5%
RUN	502,757,100	84.0%
SEDG	237,212,300	-7.8%
SPWR	538,024,300	-44.6%
TERP	151,434,700	-5.0%
VSLR	136,539,100	-3.5%
![image](https://user-images.githubusercontent.com/118132063/205224029-a5169d32-f63a-4dd7-9cf5-de4936a94434.png)
As show in the two date sets above, we can see that only two out of twelve stocks had a positive return, while in 2017 eleven out of twelve stocks had a positive return.

#
**Execution Time**
![VBA_Challenge_2017](https://user-images.githubusercontent.com/118132063/205224693-8b5946d3-05c9-4e77-8e03-a528db043b21.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/118132063/205224717-d7356779-5de6-4a08-aaca-5ceec4e3a873.png)
The refactoring execution times.

# 
**Advantages of refactoring code**
The advantages of the refactoring code is that’s it’s much more efficient  and much easier to understand. 
