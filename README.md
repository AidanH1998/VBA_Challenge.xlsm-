# VBA_Challenge.xlsm-

## **What We Have To DO?**

In this challenge we are looking at the 12 stocks that we want to invest in and in order to invest in them we have to know how they are doing in the selected years, 2017 and 2018.

What we have to do is find the return of the stock in percentage form so that we can determine if the stock in that year was green or not.

## **Why Use RSA?**

So why would we Refactor Stock Analysis, that would be because the time is fast inwhich the information is being given to you, less than a sec to get the information that you need. The con of Refactor Stock Analysis is that the code would take up the time to setup and adding in the correct data point for the stock.

## **Disadavantage of RSA?**

The coding takes time, means that maintaining the code for different stocks would be necessary because you would be using this to look at different points of the Stocks

## **Advantages of RSA?**

Your using RSA for the speed at which the information is given after the code is already done, as shown in the screenshots for 2017 and 2018 the time was less than a sec.




![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/100543143/156948643-384952a9-2666-416b-82cc-eb594c84664a.png)


![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/100543143/156948648-fb1925aa-059e-4883-a213-708829ca6a57.png)

## **Result**

In the code we have to create an area where the table will be sitting at. With the table assigned we can add the data for the stocks so that we can a a visual acceptable table which everyone can look at and be satified.

    Sub DQAnalysis()
     Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = rowStart To rowEnd

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1


    End Sub
    Sub AllStocksAnalysis()
        Dim startTime As Single
        Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer


    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"


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
  

    Dim startingPrice As Single
    Dim endingPrice As Single
    

    Worksheets(yearValue).Activate


    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    

    For i = 0 To 11
        ticker = tickers(i)
        
        totalVolume = 0
    

    Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
    
        If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
    End If
    
    
        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            startingPrice = Cells(j, 6).Value
        End If
        
        
    
        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            endingPrice = Cells(j, 6).Value
        End If
        
        
    Next j
    

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i


      Worksheets("All Stocks Analysis").Activate

      Range("A3:C3").Font.Bold = True
      Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
      Range("B4:B15").NumberFormat = "#,##0"
      Range("C4:C15").NumberFormat = "0.0%"
      Columns("B").AutoFit


    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
    
    If Cells(i, 3) > 0 Then
        Cells(i, 3).Interior.Color = vbGreen
        
    ElseIf Cells(i, 3) < 0 Then
        Cells(i, 3).Interior.Color = vbRed
        
    Else
        Cells(i, 3).Interior.Color = xlNone
        
    
    End If
    Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

     End Sub

    Sub ClearWorksheet()

    Cells.Clear

    End Sub
