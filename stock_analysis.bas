Attribute VB_Name = "Module1"
Sub Stock_Analysis()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

'Declare worksheet to execute nested loops for each year(worksheet)

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent change"
    ws.Cells(1, 12).Value = "Total Stock of Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'declaring variables for finding ticker and calculating price change, percentage change and total stock volume
    
    Dim Ticker As String
    Dim lastrow As Long
    Dim lasttickerrow As Long
    Dim i As Long
    Dim j As Integer
    Dim TickerRow As Long
    TickerRow = 1
    
    Dim openprice As Double
    Dim closeprice As Double
    Dim pricechange As Double
    Dim percentchange As Double
    Dim stockvolume As Double
    openprice = 0
    closeprice = 0
    pricechange = 0
    percentchange = 0
    stockvolume = 0
    
    
    'getting the last row with a value for the sheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'opening price for the initial ticker
    openprice = ws.Cells(2, 3).Value
    
    'first for loop to create ticker symbol, calculate price change, percentage change and total volume for each ticker
    For i = 2 To lastrow
    
        'total stock price for initial ticker
        stockvolume = stockvolume + ws.Cells(i, 7).Value
        
        
        'execute loop to find ticker and populate the Ticker column
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1).Value
            TickerRow = TickerRow + 1
            ws.Cells(TickerRow, 9).Value = Ticker
            
            
            
            'calculate price change
            closeprice = ws.Cells(i, 6).Value
            'ws.Cells(TickerRow, 14).Value = closeprice
            
            pricechange = closeprice - openprice
            
            'populating price change for each ticker
            ws.Cells(TickerRow, 10).Value = pricechange
            
            
            'correcting error for dividing by 0
            If openprice > 0 Then
                percentchange = (pricechange / openprice)
                Else
                percentchange = 0
            End If
            
            
            'populating values for percentage change and total stock volume for each ticker
            ws.Cells(TickerRow, 11).Value = percentchange
            ws.Cells(TickerRow, 12).Value = stockvolume
            
            'resetting opening price and sotck volume for each ticker after the first
            openprice = ws.Cells(i + 1, 3).Value
            stockvolume = 0
            
            'conditional formats; Green if price change is more than 0, Red if not..
            If pricechange > 0 Then
                ws.Cells(TickerRow, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(TickerRow, 10).Interior.ColorIndex = 3
            End If
            
            'changing format for percentage change to 2 decimal places and % symbol
            ws.Cells(TickerRow, 11).NumberFormat = "0.00%"
            
        End If
        
        ' Find greatest increase, decrease, and volume
            greatestincrease = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(TickerRow, 11)))
            greatestdecrease = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(TickerRow, 11)))
            greatestvolume = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(TickerRow, 12)))

            ' Update summary table
            ws.Cells(2, 16).Value = Application.WorksheetFunction.Index(ws.Range(ws.Cells(2, 9), ws.Cells(TickerRow, 9)), Application.WorksheetFunction.Match(greatestincrease, ws.Range(ws.Cells(2, 11), ws.Cells(TickerRow, 11)), 0))
            ws.Cells(3, 16).Value = Application.WorksheetFunction.Index(ws.Range(ws.Cells(2, 9), ws.Cells(TickerRow, 9)), Application.WorksheetFunction.Match(greatestdecrease, ws.Range(ws.Cells(2, 11), ws.Cells(TickerRow, 11)), 0))
            ws.Cells(4, 16).Value = Application.WorksheetFunction.Index(ws.Range(ws.Cells(2, 9), ws.Cells(TickerRow, 9)), Application.WorksheetFunction.Match(greatestvolume, ws.Range(ws.Cells(2, 12), ws.Cells(TickerRow, 12)), 0))
           ws.Cells(2, 17).Value = greatestincrease
           ws.Cells(3, 17).Value = greatestdecrease
           ws.Cells(4, 17).Value = greatestvolume
        
        'formattinggreastest increase and greastest decrease cells to 2 decimal places with % symbol
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
    Next i

Next ws

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub




