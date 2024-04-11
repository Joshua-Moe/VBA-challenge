Attribute VB_Name = "Module1"
Sub module02_VBA_challenge():
    MsgBox ("Running VBS script.")
    
    ' Loop through sheets
    For Each ws In Worksheets
    
        ' Set headers
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
    
        ' Ticker "I" ; Yearly Change "J" ; Percent Change "K" ; Total Stock Volume "L"
        
        ' Declaring variables
        Dim ticker As String
        ticker = ""
        
        Dim totalStockVolume As Double ' Long debug
        totalStockVolume = 0
        
        Dim openingStockPrice As Double
        openingStockPrice = 0
        
        Dim closingStockPrice As Double
        closingStockPrice = 0
    
    
        Dim yearlyChange As Double
        yearlyChange = 0
        Dim percentChange As Double
        percentChange = 0
        
        Dim summaryRow As Integer
        summaryRow = 2
    
        ' get length of rows
        Dim rowCount As Long
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        ' Loop through rows
        For Row = 2 To rowCount
            If ticker = "" Then
                ' First instance of a new ticker
                ticker = ws.Cells(Row, "A").Value
                openingStockPrice = ws.Cells(Row, "C").Value
                totalStockVolume = ws.Cells(Row, "G").Value
                
            ElseIf ws.Cells(Row, 1).Value = ws.Cells(Row + 1, 1).Value Then
                totalStockVolume = totalStockVolume + ws.Cells(Row, "G").Value
                
            Else
                ' End of ticker block group
                closingStockPrice = ws.Cells(Row, "F").Value
                yearlyChange = closingStockPrice - openingStockPrice
                percentChange = (closingStockPrice - openingStockPrice) / openingStockPrice
                totalStockVolume = totalStockVolume + ws.Cells(Row, "G").Value
                
                ' Update Summary section
                ws.Cells(summaryRow, "I").Value = ticker
                ws.Cells(summaryRow, "J").Value = yearlyChange
                ws.Cells(summaryRow, "K").Value = percentChange
                ws.Cells(summaryRow, "L").Value = totalStockVolume
                
                ' reset variables
                ticker = ""
                totalStockVolume = 0
                openingStockPrice = 0
                closingStockPrice = 0
                yearlyChange = 0
                percentChange = 0
                
                ' increment summaryRow
                summaryRow = summaryRow + 1
                
            End If
            
        Next Row
        
        ' Finding the "Greatest % increase", "Greatest % decrease", and
        ' "Greatest total volume"
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % increase"
        ws.Cells(3, "O").Value = "Greatest % decrease"
        ws.Cells(4, "O").Value = "Greatest total volume"
        
    
        
        ' "Greatest % increase"
        Dim greatestPecIncTicker As String
        Dim greatestPecIncValue As Double
        greatestPecIncValue = 0
        For Row = 2 To ws.Cells(Rows.Count, "K").End(xlUp).Row
            If ws.Cells(Row, "K").Value > greatestPecIncValue Then
                greatestPecIncValue = ws.Cells(Row, "K").Value
                greatestPecIncTicker = ws.Cells(Row, "I").Value
            End If
        Next Row
        ' Set values
        ws.Cells(2, "P").Value = greatestPecIncTicker
        ws.Cells(2, "Q").Value = greatestPecIncValue
        
        
        ' "Greatest % decrease"
        Dim greatestPecDecTicker As String
        Dim greatestPecDecValue As Double
        greatestPecDecValue = 0
        For Row = 2 To ws.Cells(Rows.Count, "K").End(xlUp).Row
            If ws.Cells(Row, "K").Value < greatestPecDecValue Then
                greatestPecDecValue = ws.Cells(Row, "K").Value
                greatestPecDecTicker = ws.Cells(Row, "I").Value
            End If
        Next Row
        ' Set values
        ws.Cells(3, "P").Value = greatestPecDecTicker
        ws.Cells(3, "Q").Value = greatestPecDecValue
        
        
        ' "Greatest total volume"
        Dim greatestTotVolTicker As String
        Dim greatestTotVolValue As Double
        greatestTotVolValue = 0
        For Row = 2 To ws.Cells(Rows.Count, "L").End(xlUp).Row
            If ws.Cells(Row, "L").Value > greatestTotVolValue Then
                greatestTotVolValue = ws.Cells(Row, "L").Value
                greatestTotVolTicker = ws.Cells(Row, "I").Value
            End If
        Next Row
        ' Set values
        ws.Cells(4, "P").Value = greatestTotVolTicker
        ws.Cells(4, "Q").Value = greatestTotVolValue
        
        
    Next

End Sub
