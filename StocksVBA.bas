Attribute VB_Name = "Stocks"
Sub stockTicker():
     Dim ws As Worksheet
     
     For Each ws In Worksheets
     
        'declare variables
        Dim i As Long
        Dim lastrow As Long
        Dim Ticker As String
        Dim closingprice As Double
        Dim openingprice As Double
        Dim stockvolume As LongLong
        Dim start As Double
        Dim greatestdecrese, greatestincrease, greatestvolume As Double
    
    
        'create tables
        ws.Cells(1, "J").Value = "Ticker"
        ws.Cells(1, "K").Value = "Yearly Change"
        ws.Cells(1, "L").Value = "% Change"
        ws.Cells(1, "M").Value = "Total Stock Volume"
        
        ws.Cells(2, "P") = "Greatest Percent Decrease"
        ws.Cells(3, "P") = "Greatest Percent Increase"
        ws.Cells(4, "P") = "Greatest Volume"
        ws.Cells(1, "Q") = "Ticker"
        ws.Cells(1, "R") = "Values"
    
        'Define price
        openingprice = ws.Cells(2, "C").Value
        
        start = 2
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        stockvolume = 0
        
        
        For i = 2 To lastrow
        
        stockvolume = stockvolume + ws.Cells(i, "G").Value
        Ticker = ws.Cells(i, "A").Value
    
                If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
                    ws.Cells(start, "J").Value = Ticker
                    
                    'ticker close price
                    closingprice = ws.Cells(i, "F").Value
                    
                    'calculate yearly change
                    ws.Cells(start, "K").Value = closingprice - openingprice
                    
                    'calculate percentages
                    If openingprice <> 0 Then
                        ws.Cells(start, "L").Value = FormatPercent((closingprice - openingprice) / openingprice, 2)
                        Else
                        ws.Cells(start, "L").Value = Null
                        End If
                    
                    'enter ticker volume
                    ws.Cells(start, "M").Value = stockvolume
                    
                    'Format the colors
                    If ws.Cells(start, "L").Value > 0 Then
                        ws.Cells(start, "L").Interior.ColorIndex = 4
                        Else
                        ws.Cells(start, "L").Interior.ColorIndex = 3
                        End If
                        
                    'Reset the stock volume and add one to the starting postition
                    stockvolume = 0
                    start = start + 1
                    
                    'Move to the next opening price
                    openingprice = ws.Cells(i + 1, 3).Value
                    
                'find the greatest and least percent change
                greatestdecrease = WorksheetFunction.Min(ws.Range("L2:L" & lastrow))
                greatestincrease = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
                
                'find the largest volume
                greatestvolume = WorksheetFunction.Max(ws.Range("M2:M" & lastrow))
                
                'assign the values above to greatest values table
                ws.Cells(2, "R").Value = FormatPercent(greatestdecrease)
                ws.Cells(3, "R").Value = FormatPercent(greatestincrease)
                ws.Cells(4, "R").Value = greatestvolume
                
                'find the correpsponding tickers
                decreaselocation = WorksheetFunction.Match(greatestdecrease, ws.Range("L2:L" & lastrow), 0)
                increaselocation = WorksheetFunction.Match(greatestincrease, ws.Range("L2:L" & lastrow), 0)
                volumelocation = WorksheetFunction.Match(greatestvolume, ws.Range("M2:M" & lastrow), 0)
                
                'assign the tickers to the values
                ws.Cells(2, "Q") = ws.Cells(decreaselocation + 1, "J")
                ws.Cells(3, "Q") = ws.Cells(increaselocation + 1, "J")
                ws.Cells(4, "Q") = ws.Cells(volumelocation + 1, "J")
                
                End If
                
        Next i
        
    Next ws
        
    End Sub
    
