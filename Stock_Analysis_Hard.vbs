Sub Stock_Analysis_Hard()

    Dim current As Worksheet
    
    Dim ticker, newTicker As String
    Dim stockCount, newStockCount, priceChange As Double
    
    Dim rowCount, colCount As Long
    Dim wsCount As Long
    
    Dim tickerPriceDate As Date
    Dim tickerDate As Date
    Dim dateStr As String
    
    Dim i, outputRow As Integer
    
    'Getting the count of Active sheets
    wsCount = ActiveWorkbook.Worksheets.Count
    
    'Loop through each sheet
    For Each currentSheet In ActiveWorkbook.Worksheets
    
        currentSheet.Activate
        
        'MsgBox (currentSheet.Name)
                
        'Initialize the worksheet variables
        'set title
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percentage Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'set the starting output cell
        outputRow = 2
        
        'variables to store stock - year starting and year end price
        Dim stockYearStartingPrice, stockYearEndingPrice As Double
        Dim newStockYearStartingPrice, newStockYearEndingPrice As Double
        
        ticker = ""
        newTicker = ""
        stockCount = 0
        newStockCount = 0
        stockYearStartingPrice = 0
        stockYearEndingPrice = 0
        newStockYearStartingPrice = 0
        newStockYearEndingPrice = 0
        
        'Getting the non-empty row and column count
        s_rowCount = Cells(Rows.Count, 1).End(xlUp).Row
        s_colCount = Cells(1, Columns.Count).End(xlToLeft).Column
            
        'loop all rows
        For i = 2 To s_rowCount
            
            'Call printTicker(CStr(ticker))
            
            'get the ticker name
            'hold on to this ticker till the next ticker is different
            If ticker = "" And stockCount = 0 Then
                ticker = Cells(i, 1).Value
                stockCount = Cells(i, 7).Value
                stockYearStartingPrice = Cells(i, 3).Value
                stockYearEndingPrice = Cells(i, 6).Value
            Else
                newTicker = Cells(i, 1).Value
                newStockCount = Cells(i, 7).Value
                newStockYearStartingPrice = Cells(i, 3).Value
                newStockYearEndingPrice = Cells(i, 6).Value
                
                If ticker = newTicker Then
                    
                    stockCount = stockCount + newStockCount
                    stockYearEndingPrice = newStockYearEndingPrice
                    'if the stock year start price was zero then keep changing it till we get a non-zero
                    If stockYearStartingPrice = 0 Then
                        stockYearStartingPrice = newStockYearStartingPrice
                    End If
                
                ElseIf (ticker <> newTicker) Then
                                       
                    'print to excel
                    Call PrintToSpreadsheet(CStr(ticker), CDbl(stockYearStartingPrice), CDbl(stockYearEndingPrice), CDbl(stockCount), CInt(outputRow))
                    
                    'update variables
                    stockYearStartingPrice = newStockYearStartingPrice
                    stockYearEndingPrice = newStockYearEndingPrice
                    stockCount = newStockCount
                    ticker = newTicker
                    
                    'reset variables for next loop
                    outputRow = outputRow + 1
                    newTicker = ""
                
                End If
                
                If i = s_rowCount Then
                
                    'print to excel
                    Call PrintToSpreadsheet(CStr(ticker), CDbl(stockYearStartingPrice), CDbl(stockYearEndingPrice), CDbl(stockCount), CInt(outputRow))
                    
                    'no need to update variables
                    stockYearStartingPrice = newStockYearStartingPrice
                    stockYearEndingPrice = newStockYearEndingPrice
                    stockCount = newStockCount
                    ticker = newTicker
                    
                    'reset variables for next loop
                    outputRow = outputRow + 1
                    newTicker = ""
                
                End If
                    
            End If
            
        Next i
        
        'Perform greatest and lowest % change per sheet
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        
        'Highest
        Cells(2, 16).Value = WorksheetFunction.Max(Range("K:K"))
        Cells(2, 16).NumberFormat = "0.00%"
        Cells(2, 15).Value = Cells(WorksheetFunction.Match(CDbl(Cells(2, 16).Value), Range("K:K"), 0), 9).Value
        'Lowest
        Cells(3, 16).Value = WorksheetFunction.Min(Range("K:K"))
        Cells(3, 16).NumberFormat = "0.00%"
        Cells(3, 15).Value = Cells(WorksheetFunction.Match(CDbl(Cells(3, 16).Value), Range("K:K"), 0), 9).Value
        'Greatest total volume
        Cells(4, 16).Value = WorksheetFunction.Max(Range("L:L"))
        Cells(4, 15).Value = Cells(WorksheetFunction.Match(CDbl(Cells(4, 16).Value), Range("L:L"), 0), 9).Value
        
        'find the row for the min and max to get the ticker
        'MsgBox ("Val= " & CDbl(Cells(4, 16).Value))
        
        'rowNum = WorksheetFunction.Match(CDbl(Cells(4, 16).Value), Range("L:L"), 0)
        'myTicker = Cells(rowNum, 9).Value
        
        'MsgBox ("rowNum = " & CDbl(rowNum))
        'MsgBox ("myTicker = " & CStr(myTicker))
        
    Next

End Sub

Sub PrintToSpreadsheet(ticker As String, startingPrice As Double, endingPrice As Double, stockCount As Double, printRow As Integer)

    'MsgBox ("Inside PrintToSpreadsheet")

    'print to excel
    Cells(printRow, 9).Value = ticker
    
    'Yearly Change
    Cells(printRow, 10).Value = endingPrice - startingPrice
    If Cells(printRow, 10).Value < 0 Then
        'Cell color set to Red
        Cells(printRow, 10).Interior.ColorIndex = 3
    Else
        'Cell color set to Green
        Cells(printRow, 10).Interior.ColorIndex = 4
    End If
    
    ' Percentage Change
    If startingPrice <> 0 Then
        Cells(printRow, 11).Value = (endingPrice - startingPrice) / startingPrice
    Else
        'MsgBox ("startingPrice = " & CDbl(spartingPrice))
        Cells(printRow, 11).Value = 0
    End If
    Cells(printRow, 11).NumberFormat = "0.00%"
       
    'print stock count
    Cells(printRow, 12).Value = stockCount

End Sub



